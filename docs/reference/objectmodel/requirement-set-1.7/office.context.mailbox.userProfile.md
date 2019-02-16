---
title: Office.context.mailbox.userProfile-ensemble de conditions requises 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: fb55d11fd46a9957dab124514ef3bfe5a7c138eb
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067866"
---
# <a name="userprofile"></a>userProfile

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a>[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [accountType](#accounttype-string) | Membre |
| [displayName](#displayname-string) | Membre |
| [emailAddress](#emailaddress-string) | Membre |
| [timeZone](#timezone-string) | Membre |

### <a name="members"></a>Members

####  <a name="accounttype-string"></a>accountType :String

> [!NOTE]
> Actuellement, ce membre est uniquement pris en charge par Outlook 2016 pour Mac (Build 16.9.1212 ou version ultérieure).

Obtient le type de compte de l’utilisateur associé à la boîte aux lettres. Les valeurs possibles sont répertoriées dans le tableau suivant.

| Valeur | Description |
|-------|-------------|
| `enterprise` | La boîte aux lettres se trouve sur un serveur Exchange local. |
| `gmail` | La boîte aux lettres est associée à un compte Gmail. |
| `office365` | La boîte aux lettres est associée à un compte professionnel ou scolaire Office 365. |
| `outlookCom` | La boîte aux lettres est associée à un compte Outlook.com personnel. |

##### <a name="type"></a>Type

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a>displayName :String

Obtient le nom d’affichage de l’utilisateur.

##### <a name="type"></a>Type

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a>emailAddress :String

Obtient l’adresse de messagerie SMTP de l’utilisateur.

##### <a name="type"></a>Type

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a>timeZone :String

Obtient le fuseau horaire par défaut de l’utilisateur.

##### <a name="type"></a>Type

*   Chaîne

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
