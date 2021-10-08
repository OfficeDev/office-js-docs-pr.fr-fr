---
title: Outlook prévisualisation de l’API du add-in
description: Fonctionnalités et API actuellement en prévisualisation pour Outlook de recherche.
ms.date: 10/05/2021
ms.localizationpriority: medium
ms.openlocfilehash: 7e377725bc61b33d8b67729cf041a12255e3dba5
ms.sourcegitcommit: efd0966f6400c8e685017ce0c8c016a2cbab0d5c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/08/2021
ms.locfileid: "60237188"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Outlook prévisualisation de l’API du add-in

Le sous-ensemble d’API de Outlook de l’API JavaScript Office inclut des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un Outlook.

> [!IMPORTANT]
> Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> Vous pourrez peut-être prévisualiser les fonctionnalités dans Outlook sur le web en configurant la version [ciblée sur votre Microsoft 365 client.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center) Cette page indique « Configurer l’accès en aperçu » pour les fonctionnalités applicables.
>
> Pour d’autres fonctionnalités, vous pouvez demander l’accès aux bits d’aperçu pour Outlook sur le web à l’aide de votre compte Microsoft 365 en complétant et en envoyant [ce formulaire.](https://aka.ms/OWAPreview) « Demander l’accès en prévisualisation » est indiqué sur ces fonctionnalités.

L’ensemble de conditions requises de prévisualisation inclut toutes les fonctionnalités de l’ensemble de conditions [requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a>Activation de complément sur des éléments protégés par la Gestion des droits de l’information (IRM)

Les add-ins peuvent désormais être activés sur les éléments protégés par IRM. Pour activer cette fonctionnalité, un administrateur client doit activer le droit d’utilisation en paramètres de stratégie personnalisée Autoriser l’accès par programme `OBJMODEL` dans Office.  Pour plus [d’informations, voir droits d’utilisation et descriptions.](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions)

**Disponible dans**: Outlook sur Windows, à partir de la build 13229.10000 (connectée à Microsoft 365 abonnement)

<br>

---

---

### <a name="additional-calendar-properties"></a>Propriétés de calendrier supplémentaires

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée d’un rendez-vous en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente la sensibilité d’un rendez-vous en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office.context.mailbox.item.isAllDayEvent](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui représente si un rendez-vous est un événement d’une journée.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office.context.mailbox.item.sensitivity](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui représente la sensibilité d’un rendez-vous.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums.AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

Ajout d’une nouvelle enum `AppointmentSensitivityType` qui représente les options de sensibilité disponibles sur un rendez-vous.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

<br>

---

---

### <a name="delay-delivery-time"></a>Délai de remise

#### <a name="officecontextmailboxitemdelaydeliverytime"></a>[Office.context.mailbox.item.delayDeliveryTime](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui renvoie un objet qui vous permet de gérer la date et l’heure de remise d’un message en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

#### <a name="officedelaydeliverytime"></a>[Office. DelayDeliveryTime](/javascript/api/outlook/office.delaydeliverytime?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui vous permet de gérer la date et l’heure de remise d’un message en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

<br>

---

---

### <a name="event-based-activation"></a>Activation basée sur un événement

Cette fonctionnalité a été publiée dans [l’ensemble de conditions requises 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md). Toutefois, des événements supplémentaires sont désormais disponibles en prévisualisation. Pour en savoir plus, reportez-vous aux [événements pris en charge.](../../../outlook/autolaunch.md#supported-events)

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

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

### <a name="office-theme"></a>Thème Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true#officeTheme)

Ajout de la possibilité d’obtenir un thème Office.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true)

Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365 abonnement)

<br>

---

---

### <a name="session-data"></a>Données de session

#### <a name="officesessiondata"></a>[Office. SessionData](/javascript/api/outlook/office.sessiondata?view=outlook-js-preview&preserve-view=true)

Ajout d’un nouvel objet qui représente les données de session d’un élément.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

#### <a name="officecontextmailboxitemsessiondata"></a>[Office.context.mailbox.item.sessionData](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété pour gérer les données de session d’un élément en mode Composition.

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne)

<br>

---

---

### <a name="shared-mailboxes"></a>Boîtes aux lettres partagées

La prise en charge des fonctionnalités pour les dossiers partagés (autrement dit, l’accès délégué) a été publiée dans l’ensemble de conditions [requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md). Toutefois, la prise en charge des boîtes aux lettres partagées est désormais disponible en prévisualisation. Pour plus d’informations, consultez [Activer les dossiers partagés et les scénarios de boîte aux lettres partagées](../../../outlook/delegate-access.md).

**Disponible dans**: Outlook sur Windows (connecté à un abonnement Microsoft 365), Outlook sur le web (moderne), Outlook sur Mac

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
