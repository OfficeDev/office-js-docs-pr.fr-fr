---
title: Espace de noms Office-ensemble de conditions requises 1,2
description: Les membres d’espace de noms Office sont disponibles pour les compléments Outlook à l’aide de l’API de boîte aux lettres Set 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 0dfde315cb71642a995b4c07a1966d3dee3c0d50
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431303"
---
# <a name="office-mailbox-requirement-set-12"></a>Office (boîte aux lettres requise définie sur 1,2)

L’espace de noms Office fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office, consultez la page relative à l’[interface API commune](/javascript/api/office).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="properties"></a>Propriétés

| Propriété | Modes | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [context](office.context.md) | Composition<br>Lecture | [Context](/javascript/api/office/office.context?view=outlook-js-1.2&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

##### <a name="enumerations"></a>Énumérations

| Énumération | Modes | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [AsyncResultStatus](#asyncresultstatus-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [CoercionType](#coerciontype-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [SourceProperty](#sourceproperty-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

### <a name="namespaces"></a>Espaces de noms

[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat?view=outlook-js-1.2&preserve-view=true): inclut un certain nombre d’énumérations propres à Outlook, par exemple,,,,, `ItemType` `EntityType` `AttachmentType` `RecipientType` `ResponseType` et `ItemNotificationMessageType` .

## <a name="enumeration-details"></a>Détails de l’énumération

#### <a name="asyncresultstatus-string"></a>AsyncResultStatus : chaîne

Spécifie le résultat d’un appel asynchrone.

##### <a name="type"></a>Type

*   String

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`Succeeded`| Chaîne|L’appel a réussi.|
|`Failed`| String|L’appel n’a pas réussi.|

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
