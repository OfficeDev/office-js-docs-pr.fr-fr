---
title: Espace de noms Office-ensemble de conditions requises 1,5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 36faf4569ac58693dcc1218c42a19347816d9abd
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064676"
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

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype?view=outlook-js-1.5) : inclut les énumérations ItemType, EntityType, AttachmentType, RecipientType, ResponseType et ItemNotificationMessageType.

### <a name="members"></a>Membres

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus: chaîne

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
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

---

#### <a name="coerciontype-string"></a>CoercionType: chaîne

Indique comment forcer le type des données retournées ou définies par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Html`| Chaîne|Demande que les données soient renvoyées au format HTML.|
|`Text`| Chaîne|Demande que les données soient renvoyées au format texte.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

---

#### <a name="eventtype-string"></a>EventType: chaîne

spécifie l’événement associé à un gestionnaire d’événements.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

| Nom | Type | Description |
|---|---|---|
|`ItemChanged`| Chaîne | Un autre élément Outlook est sélectionné pour consultation pendant que le volet Office est épinglé. |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1,5 |
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture |

---

#### <a name="sourceproperty-string"></a>SourceProperty: chaîne

Spécifie la source des données renvoyées par la méthode appelée.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Body`| Chaîne|La source de données est dans le corps d’un message.|
|`Subject`| Chaîne|La source de données est dans l’objet d’un message.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|
