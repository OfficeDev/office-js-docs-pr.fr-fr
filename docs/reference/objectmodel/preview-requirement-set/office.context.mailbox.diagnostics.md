---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 19cd42f845a6d96efe38da3153acf9d54339743c
ms.sourcegitcommit: 8c5c5a1bd3fe8b90f6253d9850e9352ed0b283ee
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/19/2019
ms.locfileid: "40814443"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).diagnostics

Fournit des informations de diagnostic à un complément Outlook.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

## <a name="properties"></a>Propriétés

| Propriété | Minimale<br>niveau d’autorisation | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|---|:---:|
| [Nom-d’hôte](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview#hostname) | ReadItem | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [hostVersion](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview#hostversion) | ReadItem | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [OWAView](/javascript/api/outlook/office.diagnostics?view=outlook-js-preview#owaview) | ReadItem | Composition<br>Lecture | String | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
