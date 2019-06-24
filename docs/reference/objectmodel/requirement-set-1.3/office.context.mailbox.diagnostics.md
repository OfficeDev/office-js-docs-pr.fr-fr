---
title: Office. Context. Mailbox. Diagnostics-ensemble de conditions requises 1,3
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 7e815ba2a390a871f11dfccd394ff5d306fbc0ec
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127421"
---
# <a name="diagnostics"></a>diagnostics

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmddiagnostics"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).diagnostics

Fournit des informations de diagnostic à un complément Outlook.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

### <a name="members"></a>Members

#### <a name="hostname-string"></a>NomHôte: chaîne

Obtient une chaîne qui représente le nom de l’application hôte.

Une chaîne qui peut avoir l’une des valeurs suivantes: `Outlook`, `OutlookIOS`ou`OutlookWebApp`.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

#### <a name="hostversion-string"></a>hostVersion: chaîne

Obtient une chaîne qui représente la version de l’application hôte ou du serveur Exchange Server.

Si le complément de messagerie est en cours d’exécution sur le client de bureau Outlook ou `hostVersion` sur iOS, la propriété renvoie la version de l’application hôte, Outlook. Dans Outlook sur le Web, la propriété renvoie la version du serveur Exchange. Exemple : la chaîne `15.0.468.0`.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

#### <a name="owaview-string"></a>OWAView: chaîne

Obtient une valeur de type String qui représente l’affichage actuel d’Outlook sur le Web.

La chaîne renvoyée peut avoir une des valeurs suivantes : `OneColumn`, `TwoColumns`, ou `ThreeColumns`.

Si l’application hôte n’est pas Outlook sur le Web, l’accès à cette propriété génère `undefined`.

Outlook sur le Web possède trois modes d’affichage qui correspondent à la largeur de l’écran et de la fenêtre, ainsi qu’au nombre de colonnes pouvant être affichées:

*   `OneColumn`, qui est affiché lorsque l’écran est étroit. Outlook sur le Web utilise cette disposition sur une seule colonne sur la totalité de l’écran d’un smartphone.
*   `TwoColumns`, qui est affiché lorsque l’écran est plus large. Outlook sur le Web utilise cet affichage sur la plupart des tablettes.
*   `ThreeColumns`, qui est affiché lorsque l’écran est large. Par exemple, Outlook sur le Web utilise cet affichage dans une fenêtre plein écran sur un ordinateur de bureau.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|
