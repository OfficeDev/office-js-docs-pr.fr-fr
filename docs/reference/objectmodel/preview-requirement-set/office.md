---
title: Espace de noms Office – ensemble de conditions requises
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 7effc930d196aa009c3c779b702e082ae388fada
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838514"
---
# <a name="office"></a>Office

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [AsyncResultStatus](#asyncresultstatus-string) | Member |
| [CoercionType](#coerciontype-string) | Member |
| [EventType](#eventtype-string) | Member |
| [SourceProperty](#sourceproperty-string) | Membre |

### <a name="namespaces"></a>Espaces de noms

[context](office.context.md) : fournit des interfaces partagées à partir de l’espace de noms de contexte de l’API pour les compléments Office à utiliser dans l’API du complément Outlook.

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.

### <a name="members"></a>Membres

####  <a name="asyncresultstatus-string"></a>AsyncResultStatus :String

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
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

---
---

####  <a name="coerciontype-string"></a>CoercionType :String

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| String|Demande que les données soient renvoyées au format HTML.|
|`Text`| Chaîne|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

---
---

####  <a name="eventtype-string"></a>EventType :String

spécifie l’événement associé à un gestionnaire d’événements.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

| Nom | Type | Description | Ensemble de conditions requises minimales |
|---|---|---|---|
|`AppointmentTimeChanged`| String | La date ou l’heure de la série ou du rendez-vous sélectionné a été modifiée. | 1.7 |
|`AttachmentsChanged`| String | Une pièce jointe a été ajoutée à l’élément ou supprimée de celui-ci. | Aperçu |
|`EnhancedLocationsChanged`| String | L'emplacement du rendez-vous sélectionné a changé. | Aperçu |
|`ItemChanged`| String | Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. | 1,5 |
|`OfficeThemeChanged`| Chaîne | Le thème Office de la boîte aux lettres a été modifié. | Aperçu |
|`RecipientsChanged`| String | La liste des destinataires de l’élément sélectionné ou du lieu de rendez-vous a été modifié. | 1.7 |
|`RecurrenceChanged`| Chaîne | La périodicité de la série sélectionnée a été modifiée. | 1.7 |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture |

---
---

####  <a name="sourceproperty-string"></a>SourceProperty :String

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
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|
