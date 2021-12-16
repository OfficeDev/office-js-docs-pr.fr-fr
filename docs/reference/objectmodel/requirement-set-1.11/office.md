---
title: Office de noms - ensemble de conditions requises 1.11
description: Office’espace de noms disponible pour les Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.11.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: 9b45b2afaebb0edafc41641ddc9da7bbb0de2734
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681760"
---
# <a name="office-mailbox-requirement-set-111"></a>Office (ensemble de conditions requises de boîte aux lettres 1.11)

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office?view=outlook-js-1.11&preserve-view=true).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

## <a name="properties"></a>Propriétés

| Propriété | Modes | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [context](office.context.md) | Composition<br>Lecture | [Context](/javascript/api/office/office.context?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="enumerations"></a>Énumérations

| Énumération | Modes | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [EventType](#eventtype-string) | Composition<br>Lecture | Chaîne | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [SourceProperty](#sourceproperty-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="namespaces"></a>Espaces de noms

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.11&preserve-view=true): inclut un certain nombre d’Outlook spécifiques à l’utilisateur, par exemple, `ItemType` , , , et `EntityType` `AttachmentType` `RecipientType` `ResponseType` `ItemNotificationMessageType` .

## <a name="enumeration-details"></a>Détails de l’éumération

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: String

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| String|L’appel a réussi.|
|`Failed`| String|L’appel n’a pas réussi.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

<br>

---
---

#### <a name="coerciontype-string"></a>CoercionType: String

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés

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

#### <a name="eventtype-string"></a>EventType: String

spécifie l’événement associé à un gestionnaire d’événements.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés

| Nom | Type | Description | Ensemble de conditions requises minimales |
|---|---|---|:---:|
|`AppointmentTimeChanged`| Chaîne | La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée. | 1.7 |
|`AttachmentsChanged`| Chaîne | Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci. | 1.8 |
|`EnhancedLocationsChanged`| Chaîne | L’emplacement du rendez-vous sélectionné a changé. | 1.8 |
|`ItemChanged`| Chaîne | Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | 1,5 |
|`OfficeThemeChanged`| Chaîne | Le thème Office de la boîte aux lettres a été modifié. | 1.10 |
|`RecipientsChanged`| Chaîne | La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié. | 1.7 |
|`RecurrenceChanged`| Chaîne | La périodicité de la série sélectionnée a été modifiée. | 1.7 |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5 |
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

<br>

---
---

#### <a name="sourceproperty-string"></a>SourceProperty: String

Spécifie la source des données renvoyées par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés

|Nom| Type| Description|
|---|---|---|
|`Body`| Chaîne|La source de données est dans le corps d’un message.|
|`Subject`| String|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|