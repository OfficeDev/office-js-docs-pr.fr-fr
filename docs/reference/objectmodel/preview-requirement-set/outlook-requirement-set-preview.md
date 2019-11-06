---
title: Ensemble de conditions requises de l’API du complément Outlook (aperçu)
description: ''
ms.date: 11/05/2019
localization_priority: Priority
ms.openlocfilehash: f52edea26eba6c10b24f14feb1e86d811642798b
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001620"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a>Ensemble de conditions requises de l’API du complément Outlook (aperçu)

Le sous-ensemble de l’API pour le complément Outlook de l’interface API JavaScript pour Office comprend des objets, des méthodes, des propriétés et des événements à utiliser dans un complément Outlook.

> [!IMPORTANT]
> Cette documentation a trait à un [ensemble de conditions requises](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets) en **préversion**. Ces conditions n’ont pas encore été toutes implémentées, par conséquent les clients ne pourront pas demander une aide précise concernant ces conditions. Vous ne devez pas spécifier cet ensemble de conditions dans le manifeste de votre complément.

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

L’ensemble de conditions requises présenté en aperçu comprend toutes les fonctionnalités de l’[ensemble de conditions requises 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).

## <a name="features-in-preview"></a>Fonctionnalités (aperçu) :

Les fonctionnalités suivantes sont disponibles en aperçu.

### <a name="integration-with-actionable-messages"></a>Intégration avec les messages actionnables

#### <a name="officecontextmailboxitemgetinitializationcontextasyncofficecontextmailboxitemmdgetinitializationcontextasyncoptions-callback"></a>[Office.context.mailbox.item.getInitializationContextAsync](office.context.mailbox.item.md#getinitializationcontextasyncoptions-callback)

Ajout d’une nouvelle fonction qui renvoie les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur le web (classique)

<br>

---

---

### <a name="office-theme"></a>Thème Office

#### <a name="officecontextofficethemejavascriptapiofficeofficecontextofficetheme"></a>[Office.context.officeTheme](/javascript/api/office/office.context#officetheme)

Ajout de la possibilité d’obtenir un thème Office.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

#### <a name="officeeventtypeofficethemechangedjavascriptapiofficeofficeeventtype"></a>[Office.EventType.OfficeThemeChanged](/javascript/api/office/office.eventtype)

Ajout de l’événement `OfficeThemeChanged` à `Mailbox`.

**Disponible dans** : Outlook sur Windows (connecté à l’abonnement Office 365)

<br>

---

### <a name="sso"></a>Authentification unique

#### <a name="officeruntimeauthgetaccesstokenofficedevadd-insdevelopsso-in-office-add-inssso-api-reference"></a>[OfficeRuntime.auth.getAccessToken](/office/dev/add-ins/develop/sso-in-office-add-ins#sso-api-reference)

Ajout d’un accès à `getAccessToken`, qui permet aux compléments d’[obtenir un jeton d’accès](/outlook/add-ins/authenticate-a-user-with-an-sso-token) pour l’API Microsoft Graph.

**Disponible dans** : Outlook sur Windows (connecté à Office 365), Outlook sur Mac (connecté à Office 365), Outlook sur le web (moderne), Outlook sur le web (classique)

## <a name="see-also"></a>Voir aussi

- [Compléments Outlook](/outlook/add-ins/)
- [Exemples de code pour les compléments Outlook](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [Prise en main](/outlook/add-ins/quick-start)
- [Ensembles de conditions requises et clients pris en charge](../../requirement-sets/outlook-api-requirement-sets.md)
