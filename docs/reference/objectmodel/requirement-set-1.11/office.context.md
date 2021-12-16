---
title: Office.context - ensemble de conditions requises 1.11
description: Office. Membres d’objet de contexte disponibles pour Outlook à l’aide de l’ensemble de conditions requises de l’API de boîte aux lettres 1.11.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: ee1277645afe17da5a4b547670ffe3c1d28b43e8
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681770"
---
# <a name="context-mailbox-requirement-set-111"></a>context (ensemble de conditions requises de boîte aux lettres 1.11)

### <a name="officecontext"></a>[Office](office.md).context

Office.context fournit des interfaces partagées qui sont utilisées par les modules de Office applications. Cette liste ne documente que les interfaces utilisées par les Outlook les autres. Pour obtenir une liste complète de l’espace Office.context, voir la référence [Office.context dans l’API commune.](/javascript/api/office/office.context?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

## <a name="properties"></a>Propriétés

| Propriété | Modes | Type de retour | Minimum<br>ensemble de conditions requises |
|---|---|---|:---:|
| [auth](#auth-auth) | Composition<br>Lecture | [Auth](/javascript/api/office/office.auth?view=outlook-js-1.11&preserve-view=true) | [IdentityAPI 1.3](../../requirement-sets/identity-api-requirement-sets.md) |
| [contentLanguage](#contentlanguage-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [diagnostics](#diagnostics-contextinformation) | Composition<br>Lecture | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [host](#host-hosttype) | Composition<br>Lecture | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.11&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [mailbox](office.context.mailbox.md) | Composition<br>Lecture | [Boîte aux lettres](/javascript/api/outlook/office.mailbox?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [platform](#platform-platformtype) | Composition<br>Lecture | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.11&preserve-view=true) | [1.5](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [requirements](#requirements-requirementsetsupport) | Composition<br>Lecture | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Composition<br>Lecture | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Composition<br>Lecture | [UI](/javascript/api/office/office.ui?view=outlook-js-1.11&preserve-view=true) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Détails de la propriété

#### <a name="auth-auth"></a>auth: [Auth](/javascript/api/office/office.auth?view=outlook-js-1.11&preserve-view=true)

Prend en charge l' [sign-on unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’application Office d’obtenir un jeton d’accès à l’application web du module. Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.

##### <a name="type"></a>Type

*   [Auth](/javascript/api/office/office.auth?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.10|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a>contentLanguage: String

Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.

La valeur reflète le paramètre de langue d’édition actuel spécifié avec > Options d'> langue dans `contentLanguage` l’application cliente Office’édition.  

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="diagnostics-contextinformation"></a>diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.11&preserve-view=true)

Obtient des informations sur l’environnement dans lequel le module complémentaire est en cours d’exécution.

##### <a name="type"></a>Type

*   [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage: String

Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifié par l’utilisateur pour l’interface utilisateur de l’application Office client.

La valeur reflète le paramètre de langue d’affichage actuel spécifié avec > Options d'> langue dans `displayLanguage` l’application cliente Office’affichage.  

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}

// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

<br>

---
---

#### <a name="host-hosttype"></a>host: [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.11&preserve-view=true)

Obtient Office application qui héberge le module.

> [!NOTE]
> Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir l’hôte.

##### <a name="type"></a>Type

*   [HostType](/javascript/api/office/office.hosttype?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a>platform: [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.11&preserve-view=true)

Fournit la plateforme sur laquelle le module est en cours d’exécution.

> [!NOTE]
> Vous pouvez également utiliser la propriété [Office.context.diagnostics](#diagnostics-contextinformation) pour obtenir la plateforme.

##### <a name="type"></a>Type

*   [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1,5|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.11&preserve-view=true)

Fournit une méthode pour déterminer quels ensembles de conditions requises sont pris en charge sur l’application et la plateforme actuelles.

##### <a name="type"></a>Type

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.11&preserve-view=true)

Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.

L’objet vous permet de stocker et d’accéder aux données d’un module de messagerie stocké dans la boîte aux lettres d’un utilisateur, afin qu’il soit disponible pour ce dernier lorsqu’il est en cours d’exécution à partir d’un client Outlook utilisé pour accéder à cette boîte aux `RoamingSettings` lettres.

##### <a name="type"></a>Type

*   [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)| Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

<br>

---
---

#### <a name="ui-ui"></a>Interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui?view=outlook-js-1.11&preserve-view=true)

Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants d’interface utilisateur, tels que des boîtes de dialogue, dans vos Office de données.

##### <a name="type"></a>Type

*   [UI](/javascript/api/office/office.ui?view=outlook-js-1.11&preserve-view=true)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|