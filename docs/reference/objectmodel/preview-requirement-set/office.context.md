---
title: Ensemble de conditions requises pour Office. Context-preview
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 0d47ab4f136cfc51808ca509de8177bb331e3cb7
ms.sourcegitcommit: 8fb60c3a31faedaea8b51b46238eb80c590a2491
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/14/2019
ms.locfileid: "30600269"
---
# <a name="context"></a>context

### <a name="officeofficemdcontext"></a>[Office](Office.md).context

L’espace de noms Office.context fournit des interfaces partagées qui sont utilisées par des compléments dans toutes les applications Office. Cette liste documente uniquement les interfaces utilisées par des compléments Outlook. Pour obtenir une liste complète des espaces de noms Office.context, consultez la page relative à la [référence Office.context de l’interface API commune](/javascript/api/office/office.context).

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [displayLanguage](#displaylanguage-string) | Member |
| [officeTheme](#officetheme-object) | Member |
| [roamingSettings](#roamingsettings-roamingsettings) | Membre |

### <a name="namespaces"></a>Espaces de noms

[Mailbox](office.context.mailbox.md): permet d'accéder au modèle d'objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le Web.

### <a name="members"></a>Membres

####  <a name="displaylanguage-string"></a>displayLanguage :String

Obtient les paramètres régionaux (langue) au format de balise de langue RFC 1766 spécifiés par l’utilisateur pour l’interface utilisateur de l’application hôte Office.

La valeur `displayLanguage` reflète le paramètre **Langue d’affichage** actuel spécifié dans **Fichier > Options > Langue** dans l’application hôte Office.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```javascript
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

####  <a name="officetheme-object"></a>officeTheme :Object

Permet d’accéder aux propriétés pour les couleurs du thème Office.

> [!NOTE]
> Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.

À l’aide des couleurs du thème Office, vous pouvez coordonner le modèle de couleurs de votre complément avec le thème Office actuel sélectionné par l’utilisateur dans **Fichier > Compte Office > Thème Office**, qui est appliqué à toutes les applications hôtes Office. Les couleurs du thème Office s’utilisent avec les compléments de messagerie et du volet Office.

##### <a name="type"></a>Type

*   Objet

##### <a name="properties"></a>Propriétés :

|Nom| Type| Description|
|---|---|---|
|`bodyBackgroundColor`| String|Obtient la couleur d’arrière-plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.|
|`bodyForegroundColor`| String|Obtient la couleur de premier plan du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.|
|`controlBackgroundColor`| String|Obtient la couleur d’arrière-plan du contrôle du thème Office sous la forme d’un triplet hexadécimal de couleurs.|
|`controlForegroundColor`| String|Obtient la couleur du contrôle du corps du thème Office sous la forme d’un triplet hexadécimal de couleurs.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```javascript
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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a>roamingSettings :[RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

Obtient un objet qui représente les paramètres personnalisés ou l’état d’un complément de messagerie enregistrés dans la boîte aux lettres d’un utilisateur.

L’objet `RoamingSettings` vous permet de stocker et d’accéder aux données d’un complément de messagerie conservées dans la boîte aux lettres d’un utilisateur. Ainsi, cet objet est accessible par le complément de messagerie lors de son exécution à partir d’une application cliente hôte utilisée pour accéder à la boîte aux lettres.

##### <a name="type"></a>Type

*   [RoamingSettings](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restreinte|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Rédaction ou lecture|
