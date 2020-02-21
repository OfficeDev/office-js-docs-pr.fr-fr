---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9c2c661ce870e2007bd891aee040c6b3564f7b9e
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165516"
---
# <a name="context"></a>context

### <a name="officecontext"></a>[Office](office.md).context

Office. Context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste répertorie uniquement les interfaces utilisées par les compléments Outlook. Pour obtenir la liste complète de l’espace de noms Office. Context, voir la [référence Office. Context dans l’API commune](/javascript/api/office/office.context?view=outlook-js-preview).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="properties"></a>Propriétés

| Propriété | Modes | Type de retour | Minimale<br>ensemble de conditions requises |
|---|---|---|:---:|
| [auth](#auth-auth) | Composition<br>Lecture | [Auth](/javascript/api/office/office.auth?view=outlook-js-preview) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [contentLanguage](#contentlanguage-string) | Composition<br>Lire | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [Diagnostics](#diagnostics-contextinformation) | Composition<br>Lecture | [ContextInformation](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [displayLanguage](#displaylanguage-string) | Composition<br>Lecture | Chaîne | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [hote](#host-hosttype) | Composition<br>Lire | [HostType](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [mailbox](office.context.mailbox.md) | Composition<br>Lecture | [Boîte aux lettres](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [officeTheme](#officetheme-officetheme) | Composition<br>Lecture | [OfficeTheme](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [Aperçu](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [plateforme](#platform-platformtype) | Composition<br>Lecture | [PlatformType](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [requise](#requirements-requirementsetsupport) | Composition<br>Lecture | [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [roamingSettings](#roamingsettings-roamingsettings) | Composition<br>Lecture | [RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [ui](#ui-ui) | Composition<br>Lecture | [UI](/javascript/api/office/office.ui?view=outlook-js-preview) | [1.1](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a>Détails de la propriété

#### <a name="auth-auth"></a>AUTH : [auth](/javascript/api/office/office.auth)

Prend en charge l’authentification [unique (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) en fournissant une méthode qui permet à l’hôte Office d’obtenir un jeton d’accès à l’application Web du complément. Indirectement, ceci active également le complément pour accéder aux données de Microsoft Graph de l’utilisateur sans que l’utilisateur ne doive se connecter une deuxième fois.

##### <a name="type"></a>Type

*   [Auth](/javascript/api/office/office.auth)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| Aperçu|
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

#### <a name="contentlanguage-string"></a>contentLanguage : chaîne

Obtient les paramètres régionaux (langue) spécifiés par l’utilisateur pour la modification de l’élément.

La `contentLanguage` valeur reflète le paramètre de **langue d’édition** actuel spécifié avec des options de > de **fichiers > langue** dans l’application hôte Office.

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

#### <a name="diagnostics-contextinformation"></a>Diagnostics : [ContextInformation](/javascript/api/office/office.contextinformation)

Obtient des informations sur l’environnement dans lequel le complément est en cours d’exécution.

##### <a name="type"></a>Type

*   [ContextInformation](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a>displayLanguage : chaîne

Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.

La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.

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

#### <a name="host-hosttype"></a>hôte : [HostType](/javascript/api/office/office.hosttype)

Obtient l’hôte d’application Office dans lequel le complément est en cours d’exécution.

##### <a name="type"></a>Type

*   [HostType](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a>officeTheme : [OfficeTheme](/javascript/api/office/office.officetheme)

Permet d’accéder aux propriétés pour les couleurs du thème Office.

> [!NOTE]
> Ce membre est uniquement pris en charge dans Outlook sur Windows.

L’utilisation des couleurs de thème Office vous permet de coordonner le jeu de couleurs de votre complément avec le thème Office actif sélectionné par l’utilisateur avec un **compte > le compte office > l’interface utilisateur de thème**Office, qui est appliquée à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.

##### <a name="type"></a>Type

*   [OfficeTheme](/javascript/api/office/office.officetheme)

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`bodyBackgroundColor`| Chaîne|Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.|
|`bodyForegroundColor`| Chaîne|Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.|
|`controlBackgroundColor`| String|Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.|
|`controlForegroundColor`| String|Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| Aperçu|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtype"></a>plateforme : [PlatformType](/javascript/api/office/office.platformtype)

Fournit la plateforme sur laquelle le complément est en cours d’exécution.

##### <a name="type"></a>Type

*   [PlatformType](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a>Configuration requise : [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

Fournit une méthode permettant de déterminer quels ensembles de conditions requises sont pris en charge sur l’hôte et la plateforme actuels.

##### <a name="type"></a>Type

*   [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)

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

#### <a name="roamingsettings-roamingsettings"></a>roamingSettings : [roamingSettings](/javascript/api/outlook/office.roamingsettings)

Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.

L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.

##### <a name="type"></a>Type

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Niveau d’autorisation minimal](../../../outlook/understanding-outlook-add-in-permissions.md)| Restreinte|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|

<br>

---
---

#### <a name="ui-ui"></a>interface utilisateur : [interface utilisateur](/javascript/api/office/office.ui)

Fournit des objets et des méthodes que vous pouvez utiliser pour créer et manipuler des composants de l’interface utilisateur, tels que des boîtes de dialogue, dans vos compléments Office.

##### <a name="type"></a>Type

*   [UI](/javascript/api/office/office.ui)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](../../requirement-sets/outlook-api-requirement-sets.md)| 1.1|
|[Mode Outlook applicable](../../../outlook/outlook-add-ins-overview.md#extension-points)| Rédaction ou lecture|
