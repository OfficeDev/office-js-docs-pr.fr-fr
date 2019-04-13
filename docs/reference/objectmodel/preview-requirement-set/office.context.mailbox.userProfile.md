---
title: Office. Context. Mailbox. userProfile-aperçu de l'ensemble de conditions requises
description: ''
ms.date: 04/12/2019
localization_priority: Normal
ms.openlocfilehash: 29111314f16bb9c6518b350254a3036ffa125796
ms.sourcegitcommit: 95ed6dfbfa680dbb40ff9757020fa7e5be4760b6
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 04/13/2019
ms.locfileid: "31838437"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [accountType](#accounttype-string) | Member |
| [displayName](#displayname-string) | Member |
| [emailAddress](#emailaddress-string) | Member |
| [timeZone](#timezone-string) | Membre |

### <a name="members"></a>Membres

####  <a name="accounttype-string"></a>accountType: chaîne

> [!NOTE]
> Actuellement, ce membre est uniquement pris en charge dans Outlook 2016 ou version ultérieure pour Mac (Build 16.9.1212 ou version ultérieure).

Obtient le type de compte de l'utilisateur associé à la boîte aux lettres. Les valeurs possibles sont répertoriées dans le tableau suivant.

| Valeur | Description |
|-------|-------------|
| `enterprise` | La boîte aux lettres se trouve sur un serveur Exchange local. |
| `gmail` | La boîte aux lettres est associée à un compte gmail. |
| `office365` | La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365. |
| `outlookCom` | La boîte aux lettres est associée à un compte Outlook.com personnel. |

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

####  <a name="displayname-string"></a>displayName :String

Obtient le nom d’affichage de l’utilisateur.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

####  <a name="emailaddress-string"></a>emailAddress :String

Obtient l’adresse de messagerie SMTP de l’utilisateur.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

####  <a name="timezone-string"></a>timeZone :String

Obtient le fuseau horaire par défaut de l’utilisateur.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
