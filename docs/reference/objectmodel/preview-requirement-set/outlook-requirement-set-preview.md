---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: Les fonctionnalités et les API qui sont actuellement en préversion pour les compléments Outlook.
ms.date: 04/10/2020
localization_priority: Normal
ms.openlocfilehash: 94104a9fcb239d991d585abcebdd07bcab6e315f
ms.sourcegitcommit: 3355c6bd64ecb45cea4c0d319053397f11bc9834
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/22/2020
ms.locfileid: "43744865"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble d’API de complément Outlook de l’API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements que vous pouvez utiliser dans un complément Outlook.

> [!IMPORTANT]
> Cette documentation a trait à un [ensemble de conditions requises](../../requirement-sets/outlook-api-requirement-sets.md) en **préversion**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

### <a name="additional-calendar-properties"></a>Propriétés de calendrier supplémentaires

#### <a name="isalldayevent"></a>[IsAllDayEvent](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

Ajout d’un nouvel objet qui représente la propriété d’événement d’une journée entière d’un rendez-vous en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="sensitivity"></a>[Sensitivity](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

Ajout d’un nouvel objet qui représente le critère de diffusion d’un rendez-vous en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemisalldayevent"></a>[Office. Context. Mailbox. Item. isAllDayEvent](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui indique si un rendez-vous est un événement d’une journée entière.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemsensitivity"></a>[Office. Context. Mailbox. Item. Sensitivity](office.context.mailbox.item.md#properties)

Ajout d’une nouvelle propriété qui représente le critère de diffusion d’un rendez-vous.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[Office. MailboxEnums. AppointmentSensitivityType](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

Ajout d’une nouvelle `AppointmentSensitivityType` énumération qui représente les options de critère de diffusion disponibles sur un rendez-vous.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

<br>

---

---

### <a name="append-on-send"></a>Ajouter à l’envoi

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[Office. Context. Mailbox. Item. Body. appendOnSendAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

Ajout d’une nouvelle fonction à `Body` l’objet qui ajoute des données à la fin du corps de l’élément en mode composition.

**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne)

#### <a name="extendedpermissions"></a>[ExtendedPermissions](../../manifest/extendedpermissions.md)

Ajout d’un nouvel élément au manifeste dans lequel `AppendOnSend` l’autorisation étendue doit être incluse dans la collection des autorisations étendues.

**Disponible dans**: Outlook sur Windows (connecté à l’abonnement Office 365), Outlook sur le Web (moderne)

<br>

---

---

### <a name="integration-with-actionable-messages"></a>Intégration avec les messages actionnables

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#methods)

Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)

<br>

---

---

### <a name="mail-signature"></a>Signature de courrier électronique

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[Office. Context. Mailbox. Item. Body. setSignatureAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

Ajout d’une nouvelle fonction à `Body` l’objet qui ajoute ou remplace la signature dans le corps de l’élément en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[Office. Context. Mailbox. Item. disableClientSignatureAsync](office.context.mailbox.item.md#methods)

Ajout d’une fonction qui désactive la signature client pour la boîte aux lettres d’envoi en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[Office. Context. Mailbox. Item. getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

Ajout d’une nouvelle fonction qui obtient le type de composition d’un message en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[Office. Context. Mailbox. Item. isClientSignatureEnabledAsync](office.context.mailbox.item.md#methods)

Ajout d’une fonction qui vérifie si la signature client est activée sur l’élément en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officemailboxenumscomposetype"></a>[Office. MailboxEnums. ComposeType](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

Ajout d’une nouvelle `ComposeType` énumération disponible en mode composition.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

<br>

---

---

### <a name="office-theme"></a>Thème Office

#### <a name="officecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Ajout de la possibilité d’obtenir un thème Office.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officeeventtypeofficethemechanged"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

<br>

---

---

### <a name="online-meeting-provider-integration"></a>Intégration des fournisseurs de réunions en ligne

Prise en charge supplémentaire de l’intégration des réunions en ligne dans les compléments Outlook Mobile. Pour en savoir plus, reportez-vous à la rubrique [créer un complément Outlook Mobile pour un fournisseur de réunion en ligne](../../../outlook/online-meeting.md) .

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[Point d’extension MobileOnlineMeetingCommandSurface](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

Ajout `MobileOnlineMeetingCommandSurface` du point d’extension au manifeste. Il définit l’intégration de la réunion en ligne.

**Disponible dans**: Outlook sur Android (connecté à l’abonnement Office 365)

<br>

---

---

### <a name="sso"></a>Authentification unique

#### <a name="officeruntimeauthgetaccesstoken"></a>[OfficeRuntime.auth.getAccessToken](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](../../../outlook/authenticate-a-user-with-an-sso-token.md) pour l’API Microsoft Graph.

**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](../../../outlook/outlook-add-ins-overview.md)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](../../../quickstarts/outlook-quickstart.md)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
