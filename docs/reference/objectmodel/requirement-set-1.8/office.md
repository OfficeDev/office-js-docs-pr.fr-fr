---
title: Espace de noms Office-ensemble de conditions requises 1,8
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: c5c431f7a958f1c2a956f36e90ad0f3a205c6669
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163625"
---
# <a name="office"></a>Office

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="properties"></a>Propriétés

| Propriété | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [context](office.context.md) | Composition<br>Lecture | [Context](/javascript/api/office/office.context?view=outlook-js-1.8) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Énumérations

| Énumération | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Composition<br>Lire | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Composition<br>Lire | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Composition<br>Lire | Chaîne | [1,5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Composition<br>Lire | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Espaces de noms

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.8): inclut un certain nombre d’énumérations propres à Outlook, par exemple `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType`,,,,, et `ItemNotificationMessageType`.

## <a name="enumeration-details"></a>Détails de l’énumération

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus : chaîne

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| Chaîne|L’appel a réussi.|
|`Failed`| Chaîne|L’appel n’a pas réussi.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

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
|`Html`| Chaîne|Demande que les données soient renvoyées au format HTML.|
|`Text`| String|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

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
|`AppointmentTimeChanged`| Chaîne | La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée. | 1.7 |
|`AttachmentsChanged`| Chaîne | Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci. | 1.8 |
|`EnhancedLocationsChanged`| Chaîne | L’emplacement du rendez-vous sélectionné a changé. | 1.8 |
|`ItemChanged`| Chaîne | Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | 1,5 |
|`RecipientsChanged`| Chaîne | La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié. | 1.7 |
|`RecurrenceChanged`| Chaîne | La périodicité de la série sélectionnée a été modifiée. | 1.7 |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5 |
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture |

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
|`Body`| Chaîne|La source de données est dans le corps d’un message.|
|`Subject`| String|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|
