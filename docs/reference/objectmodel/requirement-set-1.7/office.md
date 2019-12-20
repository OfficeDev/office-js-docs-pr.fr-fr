---
title: Espace de noms Office-ensemble de conditions requises 1,7
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9bfff9c45cb157d2dcd42997a01f5ada40aecfa0
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814569"
---
# <a name="office"></a>Office

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="properties"></a>Propriétés

| Propriété | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [context](office.context.md) | Composition<br>Lecture | [Context](/javascript/api/office/office.context?view=outlook-js-1.7) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Énumérations

| Énumération | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Composition<br>Lecture | String | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Espaces de noms

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.7): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.

## <a name="enumeration-details"></a>Détails de l’énumération

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus : chaîne

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| String|L’appel a réussi.|
|`Failed`| String|L’appel n’a pas réussi.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType : chaîne

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| String|Demande que les données soient renvoyées au format HTML.|
|`Text`| String|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

<br>

---
---

#### <a name="eventtype-string"></a>EventType : chaîne

spécifie l’événement associé à un gestionnaire d’événements.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

| Nom | Type | Description | Ensemble de conditions requises minimales |
|---|---|---|:---:|
|`AppointmentTimeChanged`| String | La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée. | 1.7 |
|`ItemChanged`| String | Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | 1,5 |
|`RecipientsChanged`| Chaîne | La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié. | 1.7 |
|`RecurrenceChanged`| Chaîne | La périodicité de la série sélectionnée a été modifiée. | 1.7 |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5 |
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture |

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty : chaîne

Spécifie la source des données renvoyées par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Body`| String|La source de données est dans le corps d’un message.|
|`Subject`| String|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|
