---
title: Office. Context. Mailbox. Item-Preview ensemble de conditions requises
description: ''
ms.date: 11/05/2019
localization_priority: Normal
ms.openlocfilehash: a529dff046f48eff65b70813617bbb9d216dba5e
ms.sourcegitcommit: 21aa084875c9e07a300b3bbe8852b3e5dd163e1d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/06/2019
ms.locfileid: "38001627"
---
# <a name="item"></a>élément

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmditem"></a>[Office](office.md)[.context](office.context.md)[.mailbox](office.context.mailbox.md).item

L’espace de noms `item` est utilisé pour accéder au message, à la demande de réunion ou au rendez-vous actuellement sélectionné. Vous pouvez déterminer le type de l’élément `item` à l’aide de la propriété [itemType](#itemtype-officemailboxenumsitemtype).

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|Restreinte|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [attachments](#attachments-arrayattachmentdetails) | Membre |
| [bcc](#bcc-recipients) | Membre |
| [body](#body-body) | Membre |
| [catégories](#categories-categories) | Membre |
| [cc](#cc-arrayemailaddressdetailsrecipients) | Membre |
| [conversationId](#nullable-conversationid-string) | Membre |
| [dateTimeCreated](#datetimecreated-date) | Membre |
| [dateTimeModified](#datetimemodified-date) | Membre |
| [end](#end-datetime) | Membre |
| [enhancedLocation](#enhancedlocation-enhancedlocation) | Membre |
| [from](#from-emailaddressdetailsfrom) | Membre |
| [internetHeaders](#internetheaders-internetheaders) | Membre |
| [internetMessageId](#internetmessageid-string) | Membre |
| [itemClass](#itemclass-string) | Membre |
| [itemId](#nullable-itemid-string) | Membre |
| [itemType](#itemtype-officemailboxenumsitemtype) | Membre |
| [location](#location-stringlocation) | Membre |
| [normalizedSubject](#normalizedsubject-string) | Membre |
| [notificationMessages](#notificationmessages-notificationmessages) | Member |
| [optionalAttendees](#optionalattendees-arrayemailaddressdetailsrecipients) | Membre |
| [organizer](#organizer-emailaddressdetailsorganizer) | Membre |
| [recurrence](#nullable-recurrence-recurrence) | Membre |
| [requiredAttendees](#requiredattendees-arrayemailaddressdetailsrecipients) | Membre |
| [sender](#sender-emailaddressdetails) | Member |
| [seriesId](#nullable-seriesid-string) | Member |
| [start](#start-datetime) | Member |
| [subject](#subject-stringsubject) | Membre |
| [to](#to-arrayemailaddressdetailsrecipients) | Membre |
| [addFileAttachmentAsync](#addfileattachmentasyncuri-attachmentname-options-callback) | Méthode |
| [addFileAttachmentFromBase64Async](#addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback) | Méthode |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Méthode |
| [addItemAttachmentAsync](#additemattachmentasyncitemid-attachmentname-options-callback) | Méthode |
| [close](#close) | Méthode |
| [displayReplyAllForm](#displayreplyallformformdata-callback) | Méthode |
| [displayReplyForm](#displayreplyformformdata-callback) | Méthode |
| [getAllInternetHeadersAsync](#getallinternetheadersasyncoptions-callback) | Méthode |
| [getAttachmentContentAsync](#getattachmentcontentasyncattachmentid-options-callback--attachmentcontent) | Méthode |
| [getAttachmentsAsync](#getattachmentsasyncoptions-callback--arrayattachmentdetails) | Méthode |
| [getEntities](#getentities--entities) | Méthode |
| [getEntitiesByType](#getentitiesbytypeentitytype--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | Méthode |
| [getFilteredEntitiesByName](#getfilteredentitiesbynamename--nullable-arraystringcontactmeetingsuggestionphonenumbertasksuggestion) | Méthode |
| [getInitializationContextAsync](#getinitializationcontextasyncoptions-callback) | Méthode |
| [getItemIdAsync](#getitemidasyncoptions-callback) | Méthode |
| [getRegExMatches](#getregexmatches--object) | Méthode |
| [getRegExMatchesByName](#getregexmatchesbynamename--nullable-array-string-) | Méthode |
| [getSelectedDataAsync](#getselecteddataasynccoerciontype-options-callback--string) | Méthode |
| [getSelectedEntities](#getselectedentities--entities) | Méthode |
| [getSelectedRegExMatches](#getselectedregexmatches--object) | Méthode |
| [getSharedPropertiesAsync](#getsharedpropertiesasyncoptions-callback) | Méthode |
| [loadCustomPropertiesAsync](#loadcustompropertiesasynccallback-usercontext) | Méthode |
| [removeAttachmentAsync](#removeattachmentasyncattachmentid-options-callback) | Méthode |
| [removeHandlerAsync](#removehandlerasynceventtype-options-callback) | Méthode |
| [saveAsync](#saveasyncoptions-callback) | Méthode |
| [setSelectedDataAsync](#setselecteddataasyncdata-options-callback) | Méthode |

### <a name="example"></a>Exemple

L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    var subject = item.subject;
    // Continue with processing the subject of the current item,
    // which can be a message or appointment.
  });
};
```

### <a name="members"></a>Members

#### <a name="attachments-arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>attachments: Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Obtient les pièces jointes de l’élément sous la forme d’un tableau. Mode Lecture uniquement.

> [!NOTE]
> Certains types de fichiers sont bloqués par Outlook car ils présentent des problèmes de sécurité potentiels. Dans ce cas, ils ne sont pas renvoyés. Pour en savoir plus, consultez l’article [Pièces jointes bloquées dans Outlook](https://support.office.com/article/Blocked-attachments-in-Outlook-434752E1-02D3-4E90-9124-8B81E49A8519).

##### <a name="type"></a>Type

*   Array.<[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

Le code suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.

```js
var item = Office.context.mailbox.item;
var outputString = "";

if (item.attachments.length > 0) {
  for (i = 0 ; i < item.attachments.length ; i++) {
    var attachment = item.attachments[i];
    outputString += "<BR>" + i + ". Name: ";
    outputString += attachment.name;
    outputString += "<BR>ID: " + attachment.id;
    outputString += "<BR>contentType: " + attachment.contentType;
    outputString += "<BR>size: " + attachment.size;
    outputString += "<BR>attachmentType: " + attachment.attachmentType;
    outputString += "<BR>isInline: " + attachment.isInline;
  }
}

console.log(outputString);
```

<br>

---
---

#### <a name="bcc-recipientsjavascriptapioutlookofficerecipients"></a>bcc: [Recipients](/javascript/api/outlook/office.recipients)

Permet d’obtenir un objet qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne Cci (copie carbone invisible) d’un message. Mode composition uniquement.

Par défaut, la collection est limitée à 100 membres. Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.

- Obtenez 500 membres au maximum.
- Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.

##### <a name="type"></a>Type

*   [Destinataires](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="example"></a>Exemple

```js
Office.context.mailbox.item.bcc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.bcc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.bcc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfBccRecipients = asyncResult.value;
}
```

<br>

---
---

#### <a name="body-bodyjavascriptapioutlookofficebody"></a>body: [Body](/javascript/api/outlook/office.body)

Obtient un objet qui fournit des méthodes permettant de manipuler le corps d’un élément.

##### <a name="type"></a>Type

*   [Body](/javascript/api/outlook/office.body)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

Cet exemple obtient le corps du message en texte brut.

```js
Office.context.mailbox.item.body.getAsync(
  "text",
  { asyncContext: "This is passed to the callback" },
  function callback(result) {
    // Do something with the result.
  });

```

L’exemple suivant présente le paramètre de résultat transmis à la fonction de rappel.

```json
{
  "value": "TEXT of whole body (including threads below)",
  "status": "succeeded",
  "asyncContext": "This is passed to the callback"
}
```

<br>

---
---

#### <a name="categories-categoriesjavascriptapioutlookofficecategories"></a>Catégories : [catégories](/javascript/api/outlook/office.categories)

Obtient un objet qui fournit des méthodes pour la gestion des catégories de l’élément.

> [!NOTE]
> Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.

##### <a name="type"></a>Type

*   [Categories](/javascript/api/outlook/office.categories)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

Cet exemple obtient les catégories de l’élément.

```js
Office.context.mailbox.item.categories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="cc-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>cc: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Permet d’accéder aux destinataires en copie carbone (Cc) d’un message. Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.

##### <a name="read-mode"></a>Mode lecture

La propriété `cc` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **Cc** du message. Par défaut, la collection est limitée à 100 membres. Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.

```js
console.log(JSON.stringify(Office.context.mailbox.item.cc));
```

##### <a name="compose-mode"></a>Mode composition

La propriété `cc` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **Cc** du message. Par défaut, la collection est limitée à 100 membres. Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.

- Obtenez 500 membres au maximum.
- Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.

```js
Office.context.mailbox.item.cc.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.cc.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.cc.getAsync(callback);

function callback(asyncResult) {
  var arrayOfCcRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Type

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="nullable-conversationid-string"></a>(nullable) conversationId: String

Obtient l’identificateur de la conversation qui contient un message particulier.

Vous pouvez obtenir un nombre entier de cette propriété si votre application de messagerie est activée dans les formulaires de lecture ou les réponses des formulaires de composition. Si, par la suite, l’utilisateur modifie l’objet du message de réponse, lors de l’envoi de la réponse, l’ID de conversation de ce message va changer et la valeur que vous avez obtenue plus tôt ne sera plus applicable.

Cette propriété obtient une valeur null lorsqu’un élément est ajouté à un formulaire de composition. Si l’utilisateur définit la ligne Objet et enregistre l’élément, la propriété `conversationId` renvoie une valeur.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
var conversationId = Office.context.mailbox.item.conversationId;
console.log("conversationId: " + conversationId);
```

<br>

---
---

#### <a name="datetimecreated-date"></a>dateTimeCreated: Date

Obtient la date et l’heure de création d’un élément. Mode lecture uniquement.

##### <a name="type"></a>Type

*   Date

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
var dateTimeCreated = Office.context.mailbox.item.dateTimeCreated;
console.log("Date and time created: " + dateTimeCreated);
```

<br>

---
---

#### <a name="datetimemodified-date"></a>dateTimeModified: Date

Permet d’obtenir la date et l’heure de la dernière modification d’un élément. Mode lecture uniquement.

> [!NOTE]
> Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.

##### <a name="type"></a>Type

*   Date

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
var dateTimeModified = Office.context.mailbox.item.dateTimeModified;
console.log("Date and time modified: " + dateTimeModified);
```

<br>

---
---

#### <a name="end-datetimejavascriptapioutlookofficetime"></a>end: Date|[Time](/javascript/api/outlook/office.time)

Obtient ou définit la date et l’heure de fin du rendez-vous.

La propriété `end` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur de fin de la propriété à la date et à l’heure du client.

##### <a name="read-mode"></a>Mode lecture

La propriété `end` renvoie un objet `Date`.

```js
var end = Office.context.mailbox.item.end;
console.log("Appointment end: " + end);
```

##### <a name="compose-mode"></a>Mode composition

La propriété `end` renvoie un objet `Time`.

Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de fin, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.

L’exemple suivant définit l’heure de fin d’un rendez-vous en utilisant la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.

```js
var endTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.end.setAsync(endTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("End Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>Type

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="enhancedlocation-enhancedlocationjavascriptapioutlookofficeenhancedlocation"></a>enhancedLocation : [enhancedLocation](/javascript/api/outlook/office.enhancedlocation)

Obtient ou définit les emplacements d’un rendez-vous.

##### <a name="read-mode"></a>Mode lecture

La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui vous permet d’obtenir l’ensemble des emplacements (chacun représenté par un objet [LocationDetails](/javascript/api/outlook/office.locationdetails) ) associé au rendez-vous.

##### <a name="compose-mode"></a>Mode composition

La `enhancedLocation` propriété renvoie un objet [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation) qui fournit des méthodes pour obtenir, supprimer ou ajouter des emplacements sur un rendez-vous.

##### <a name="type"></a>Type

*   [EnhancedLocation](/javascript/api/outlook/office.enhancedlocation)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

L’exemple suivant obtient les emplacements actuels associés au rendez-vous.

```js
Office.context.mailbox.item.enhancedLocation.getAsync(callbackFunction);

function callbackFunction(asyncResult) {
  asyncResult.value.forEach(function (place) {
    console.log("Display name: " + place.displayName);
    console.log("Type: " + place.locationIdentifier.type);
    if (place.locationIdentifier.type === Office.MailboxEnums.LocationType.Room) {
      console.log("Email address: " + place.emailAddress);
    }
  });
}
```

<br>

---
---

#### <a name="from-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsfromjavascriptapioutlookofficefrom"></a>from : [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)|[from](/javascript/api/outlook/office.from)

Obtient l’adresse de messagerie de l’expéditeur d’un message.

Les propriétés `from` et [`sender`](#sender-emailaddressdetails) représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.

> [!NOTE]
> La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `from` est `undefined`.

##### <a name="read-mode"></a>Mode lecture

La `from` propriété renvoie un `EmailAddressDetails` objet.

```js
var from = Office.context.mailbox.item.from;
console.log("From " + from);
```

##### <a name="compose-mode"></a>Mode composition

La `from` propriété renvoie un `From` objet qui fournit une méthode pour obtenir la valeur de.

```js
Office.context.mailbox.item.from.getAsync(callback);

function callback(asyncResult) {
  var from = asyncResult.value;
}
```

##### <a name="type"></a>Type

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) | [à partir de](/javascript/api/outlook/office.from)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|||
|---|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|Composition|

<br>

---
---

#### <a name="internetheaders-internetheadersjavascriptapioutlookofficeinternetheaders"></a>internetHeaders : [internetHeaders](/javascript/api/outlook/office.internetheaders)

Obtient ou définit les en-têtes Internet personnalisés d’un message. Mode composition uniquement.

##### <a name="type"></a>Type

*   [InternetHeaders](/javascript/api/outlook/office.internetheaders)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="example"></a>Exemple

```js
Office.context.mailbox.item.internetHeaders.getAsync(["header1", "header2"], callback);

function callback(asyncResult) {
  var dictionary = asyncResult.value;
  var header1_value = dictionary["header1"];
}
```

<br>

---
---

#### <a name="internetmessageid-string"></a>internetMessageId: String

Obtient l’identificateur de message Internet d’un message électronique. Mode lecture uniquement.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
var internetMessageId = Office.context.mailbox.item.internetMessageId;
console.log("internetMessageId: " + internetMessageId);
```

<br>

---
---

#### <a name="itemclass-string"></a>itemClass: String

Obtient la classe de l’élément des services web Exchange de l’élément sélectionné. Mode lecture uniquement.

La propriété `itemClass` spécifie la classe de message de l’élément sélectionné. Les éléments suivants sont les classes de message par défaut du message ou de l’élément de rendez-vous.

|Type|Description|Classe de l’élément|
|---|---|---|
|Éléments de rendez-vous|Ce sont les éléments de calendrier de la classe de l’élément `IPM.Appointment` ou `IPM.Appointment.Occurrence`.|`IPM.Appointment`<br />`IPM.Appointment.Occurrence`|
|Éléments de message|Ces éléments incluent les messages électroniques dont la classe de message par défaut est `IPM.Note`, ainsi que les demandes de réunion, les réponses et les annulations qui utilisent `IPM.Schedule.Meeting` comme classe de message de base.|`IPM.Note`<br />`IPM.Schedule.Meeting.Request`<br />`IPM.Schedule.Meeting.Neg`<br />`IPM.Schedule.Meeting.Pos`<br />`IPM.Schedule.Meeting.Tent`<br />`IPM.Schedule.Meeting.Canceled`|

Vous pouvez créer des classes de message personnalisées qui étendent une classe de message par défaut, par exemple, une classe de message de rendez-vous personnalisée `IPM.Appointment.Contoso`.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
var itemClass = Office.context.mailbox.item.itemClass;
console.log("Item class: " + itemClass);
```

<br>

---
---

#### <a name="nullable-itemid-string"></a>(nullable) itemId: String

Obtient l' [identificateur d’élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange) pour l’élément actuel. Mode Lecture uniquement.

> [!NOTE]
> L’identificateur renvoyé par la `itemId` propriété est identique à l’identificateur d' [élément des services Web Exchange](/exchange/client-developer/exchange-web-services/ews-identifiers-in-exchange). La propriété `itemId` n’est pas identique à l’ID d’entrée Outlook ni à l’ID utilisé par l’API REST Outlook. Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api#get-the-item-id).

La propriété `itemId` n’est pas disponible en mode composition. Si l’identificateur d’un élément doit être indiqué, la méthode [`saveAsync`](#saveasyncoptions-callback) peut être utilisée pour enregistrer l’élément sur le magasin, lequel renvoie l’identificateur de l’élément dans le paramètre [`AsyncResult.value`](/javascript/api/office/office.asyncresult) dans la fonction de rappel.

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

Le code suivant vérifie la présence d’un identificateur d’élément. Si la propriété `itemId` renvoie `null` ou `undefined`, il enregistre l’élément sur le magasin et obtient l’identificateur de l’élément à partir du résultat asynchrone.

```js
var itemId = Office.context.mailbox.item.itemId;
if (itemId === null || itemId == undefined) {
  Office.context.mailbox.item.saveAsync(function(result) {
    itemId = result.value;
  });
}
```

<br>

---
---

#### <a name="itemtype-officemailboxenumsitemtypejavascriptapioutlookofficemailboxenumsitemtype"></a>itemType: [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

Obtient le type d’élément représenté par une instance.

La propriété `itemType` renvoie une des valeurs d’énumération `ItemType` indiquant si l’instance d’objet `item` est un message ou un rendez-vous.

##### <a name="type"></a>Type

*   [Office.MailboxEnums.ItemType](/javascript/api/outlook/office.mailboxenums.itemtype)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
if (Office.context.mailbox.item.itemType === Office.MailboxEnums.ItemType.Message) {
  // Do something.
} else {
  // Do something else.
}
```

<br>

---
---

#### <a name="location-stringlocationjavascriptapioutlookofficelocation"></a>location: String|[Location](/javascript/api/outlook/office.location)

Obtient ou définit le lieu d’un rendez-vous.

##### <a name="read-mode"></a>Mode lecture

La propriété `location` renvoie une chaîne contenant le lieu du rendez-vous.

```js
var location = Office.context.mailbox.item.location;
console.log("location: " + location);
```

##### <a name="compose-mode"></a>Mode composition

La propriété `location` renvoie un objet `Location` qui fournit les méthodes utilisées pour obtenir et définir le lieu du rendez-vous.

```js
var userContext = { value : 1 };
Office.context.mailbox.item.location.getAsync( { context: userContext}, callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var location = asyncResult.value;
}
```

##### <a name="type"></a>Type

*   String | [Location](/javascript/api/outlook/office.location)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="normalizedsubject-string"></a>normalizedSubject: String

Obtient l’objet d’un élément, sans les préfixes (y compris `RE:` et `FWD:`). Mode lecture uniquement.

La propriété normalizedSubject obtient l’objet de l’élément, sans les préfixes standard (par exemple, `RE:` et `FW:`) qui sont ajoutés par les programmes de messagerie électronique. Pour obtenir l’objet de l’élément avec les préfixes intacts, utilisez la propriété [`subject`](#subject-stringsubject).

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
var normalizedSubject = Office.context.mailbox.item.normalizedSubject;
console.log("Normalized subject: " + normalizedSubject);
```

<br>

---
---

#### <a name="notificationmessages-notificationmessagesjavascriptapioutlookofficenotificationmessages"></a>notificationMessages: [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

Obtient les messages de notification pour un élément.

##### <a name="type"></a>Type

*   [NotificationMessages](/javascript/api/outlook/office.notificationmessages)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
// Get all notifications.
Office.context.mailbox.item.notificationMessages.getAllAsync(
  function (asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

<br>

---
---

#### <a name="optionalattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>optionalAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Permet d’accéder aux participants facultatifs d’un événement. Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.

##### <a name="read-mode"></a>Mode Lecture

La propriété `optionalAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant facultatif à la réunion. Par défaut, la collection est limitée à 100 membres. Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.

```js
var optionalAttendees = Office.context.mailbox.item.optionalAttendees;
console.log("Optional attendees: " + JSON.stringify(optionalAttendees));
```

##### <a name="compose-mode"></a>Mode composition

La propriété `optionalAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants facultatifs d’une réunion. Par défaut, la collection est limitée à 100 membres. Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.

- Obtenez 500 membres au maximum.
- Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.

```js
Office.context.mailbox.item.optionalAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.optionalAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfOptionalAttendeesRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Type

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="organizer-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsorganizerjavascriptapioutlookofficeorganizer"></a>Organisateur : [](/javascript/api/outlook/office.emailaddressdetails)|[organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails

Obtient l’adresse de messagerie de l’organisateur d’une réunion spécifiée.

##### <a name="read-mode"></a>Mode lecture

La `organizer` propriété renvoie un objet [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) qui représente l’organisateur de la réunion.

```js
var organizerName = Office.context.mailbox.item.organizer.displayName;
var organizerAddress = Office.context.mailbox.item.organizer.emailAddress;
console.log("Organizer: " + organizerName + " (" + organizerAddress + ")");
```

##### <a name="compose-mode"></a>Mode composition

La `organizer` propriété renvoie un objet [organisateur](/javascript/api/outlook/office.organizer) qui fournit une méthode pour obtenir la valeur de l’organisateur.

```js
Office.context.mailbox.item.organizer.getAsync(
  function(asyncResult) {
    console.log(JSON.stringify(asyncResult));
  }
);
```

##### <a name="type"></a>Type

*   [](/javascript/api/outlook/office.emailaddressdetails) | [Organisateur](/javascript/api/outlook/office.organizer) EmailAddressDetails

##### <a name="requirements"></a>Configuration requise

|Conditions requises|||
|---|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|1.7|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|Composition|

<br>

---
---

#### <a name="nullable-recurrence-recurrencejavascriptapioutlookofficerecurrence"></a>(Nullable) récurrence : [périodicité](/javascript/api/outlook/office.recurrence)

Obtient ou définit la périodicité d’un rendez-vous. Obtient la périodicité d’une demande de réunion. Modes lecture et composition pour les éléments de rendez-vous. Mode lecture pour les éléments de demande de réunion.

La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) pour les demandes de réunion ou de rendez-vous périodiques si un élément est une série ou une instance dans une série. `null`est renvoyé pour les rendez-vous uniques et les demandes de réunion de rendez-vous uniques. `undefined`est renvoyée pour les messages qui ne sont pas des demandes de réunion.

> Remarque : les demandes de réunion `itemClass` ont la valeur IPM. Schedule. Meeting. Request.

> Remarque : si l’objet de périodicité `null`est, cela indique que l’objet est un rendez-vous unique ou une demande de réunion d’un seul rendez-vous et non d’une série.

##### <a name="read-mode"></a>Mode lecture

La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui représente la périodicité du rendez-vous. Elle est disponible pour les rendez-vous et les demandes de réunion.

```js
var recurrence = Office.context.mailbox.item.recurrence;
console.log("Recurrence: " + JSON.stringify(recurrence));
```

##### <a name="compose-mode"></a>Mode composition

La `recurrence` propriété renvoie un objet [Recurrence](/javascript/api/outlook/office.recurrence) qui fournit des méthodes pour gérer la périodicité des rendez-vous. Elle est disponible pour les rendez-vous.

```js
Office.context.mailbox.item.recurrence.getAsync(callback);

function callback(asyncResult) {
  var context = asyncResult.context;
  var recurrence = asyncResult.value;
  if (!recurrence) {
    console.log("One-time appointment or meeting");
  } else {
    console.log(JSON.stringify(recurrence));
  }
}

// The following example shows the results of the getAsync call that retrieves the recurrence for a series.
// NOTE: In this example, seriesTimeObject is a placeholder for the JSON representing the
// recurrence.seriesTime property. You should use the SeriesTime object's methods to get the
// recurrence date and time properties.
Recurrence = {
  "recurrenceType": "weekly",
  "recurrenceProperties": {"interval": 2, "days": ["mon","thu","fri"], "firstDayOfWeek": "sun"},
  "seriesTime": {seriesTimeObject},
  "recurrenceTimeZone": {"name": "Pacific Standard Time", "offset": -480}
}
```

##### <a name="type"></a>Type

* [Instances](/javascript/api/outlook/office.recurrence)

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="requiredattendees-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>requiredAttendees: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Permet d’accéder aux participants requis à un événement. Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.

##### <a name="read-mode"></a>Mode Lecture

La propriété `requiredAttendees` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque participant requis à la réunion. Par défaut, la collection est limitée à 100 membres. Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.

```js
var requiredAttendees = Office.context.mailbox.item.requiredAttendees;
console.log("Required attendees: " + JSON.stringify(requiredAttendees));
```

##### <a name="compose-mode"></a>Mode composition

La propriété `requiredAttendees` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les participants requis à une réunion. Par défaut, la collection est limitée à 100 membres. Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.

- Obtenez 500 membres au maximum.
- Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.

```js
Office.context.mailbox.item.requiredAttendees.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.requiredAttendees.getAsync(callback);

function callback(asyncResult) {
  var arrayOfRequiredAttendeesRecipients = asyncResult.value;
  console.log(JSON.stringify(arrayOfRequiredAttendeesRecipients));
}
```

##### <a name="type"></a>Type

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="sender-emailaddressdetailsjavascriptapioutlookofficeemailaddressdetails"></a>sender: [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

Obtient l’adresse de messagerie de l’expéditeur d’un message électronique. Mode lecture uniquement.

Les propriétés [`from`](#from-emailaddressdetailsfrom) et `sender` représentent la même personne, sauf si le message est envoyé par un délégué. Dans ce cas, la propriété `from` représente le délégant et la propriété sender représente le délégué.

> [!NOTE]
> La propriété `recipientType` de l’objet `EmailAddressDetails` dans la propriété `sender` est `undefined`.

##### <a name="type"></a>Type

*   [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
var senderName = Office.context.mailbox.item.sender.displayName;
var senderAddress = Office.context.mailbox.item.sender.emailAddress;
console.log("Sender: " + senderName + " (" + senderAddress + ")");
```

<br>

---
---

#### <a name="nullable-seriesid-string"></a>(Nullable) seriesId : chaîne

Obtient l’ID de la série à laquelle une instance appartient.

Dans Outlook sur le Web et les clients de bureau `seriesId` , le renvoie l’ID des services Web Exchange (EWS) de l’élément parent (série) auquel cet élément appartient. Toutefois, dans iOS et Android, le `seriesId` renvoie l’ID REST de l’élément parent.

> [!NOTE]
> L’identificateur renvoyé par la propriété `seriesId` est identique à celui de l’élément des services web Exchange. La `seriesId` propriété n’est pas identique aux ID Outlook utilisés par l’API REST Outlook. Avant que vous ne puissiez effectuer des appels d’API REST avec cette valeur, elle doit être convertie à l’aide de la commande [Office.context.mailbox.convertToRestId](office.context.mailbox.md#converttorestiditemid-restversion--string). Pour plus d’informations, voir [Utilisation des API REST Outlook à partir d’un complément Outlook](/outlook/add-ins/use-rest-api).

La `seriesId` propriété renvoie `null` pour les éléments qui n’ont pas d’éléments parents, tels que les rendez-vous uniques, les `undefined` éléments de série ou les demandes de réunion, et les retours pour tous les autres éléments qui ne sont pas des demandes de réunion.

##### <a name="type"></a>Type

* String

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.7|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
var seriesId = Office.context.mailbox.item.seriesId;

// The seriesId property returns null for items that do
// not have parent items (such as single appointments,
// series items, or meeting requests) and returns
// undefined for messages that are not meeting requests.
var isSeriesInstance = (seriesId != null);
console.log("SeriesId is " + seriesId + " and isSeriesInstance is " + isSeriesInstance);
```

<br>

---
---

#### <a name="start-datetimejavascriptapioutlookofficetime"></a>start: Date|[Time](/javascript/api/outlook/office.time)

Obtient ou définit la date et l’heure de début du rendez-vous.

La propriété `start` est exprimée en date et heure UTC (temps universel coordonné). Vous pouvez utiliser la méthode [`convertToLocalClientTime`](office.context.mailbox.md#converttolocalclienttimetimevalue--localclienttime) pour convertir la valeur à la date et à l’heure du client.

##### <a name="read-mode"></a>Mode lecture

La propriété `start` renvoie un objet `Date`.

```js
var start = Office.context.mailbox.item.start;
console.log("Appointment start: " + JSON.stringify(start));
```

##### <a name="compose-mode"></a>Mode composition

La propriété `start` renvoie un objet `Time`.

Quand vous utilisez la méthode [`Time.setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) pour définir l’heure de début, nous vous recommandons d’utiliser la méthode [`convertToUtcClientTime`](office.context.mailbox.md#converttoutcclienttimeinput--date) pour convertir l’heure locale du client au format UTC pour le serveur.

L’exemple suivant définit l’heure de début d’un rendez-vous en mode composition à l’aide de la méthode [`setAsync`](/javascript/api/outlook/office.time#setasync-datetime--options--callback-) de l’objet `Time`.

```js
var startTime = new Date("3/14/2015");
var options = {
  // Pass information that can be used in the callback.
  asyncContext: {verb: "Set"}
};
Office.context.mailbox.item.start.setAsync(startTime, options, function(result) {
  if (result.error) {
    console.debug(result.error);
  } else {
    // Access the asyncContext that was passed to the setAsync function.
    console.debug("Start Time " + result.asyncContext.verb);
  }
});
```

##### <a name="type"></a>Type

*   Date | [Time](/javascript/api/outlook/office.time)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="subject-stringsubjectjavascriptapioutlookofficesubject"></a>subject: String|[Subject](/javascript/api/outlook/office.subject)

Obtient ou définit la description qui apparaît dans le champ d’objet d’un élément.

La propriété `subject` obtient ou définit l’intégralité de l’objet de l’élément, tel qu’il est envoyé par le serveur de messagerie.

##### <a name="read-mode"></a>Mode lecture

La propriété `subject` renvoie une chaîne. Utilisez la propriété [`normalizedSubject`](#normalizedsubject-string) pour obtenir l’objet sans les préfixes tels que `RE:` et `FW:`.

L’exemple de code JavaScript suivant montre comment accéder à la propriété `subject` de l’élément actif dans Outlook.

```js
var subject = Office.context.mailbox.item.subject;
console.log(subject);
```

##### <a name="compose-mode"></a>Mode composition
La propriété `subject` renvoie un objet `Subject` qui fournit des méthodes pour obtenir et définir l’objet.

```js
Office.context.mailbox.item.subject.getAsync(callback);

function callback(asyncResult) {
  var subject = asyncResult.value;
  console.log(subject);
}
```

##### <a name="type"></a>Type

*   String | [Subject](/javascript/api/outlook/office.subject)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

<br>

---
---

#### <a name="to-arrayemailaddressdetailsjavascriptapioutlookofficeemailaddressdetailsrecipientsjavascriptapioutlookofficerecipients"></a>to: Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)>|[Recipients](/javascript/api/outlook/office.recipients)

Permet d’accéder aux destinataires figurant sur la ligne **À** d’un message. Le type d’objet et le niveau d’accès varient selon le mode de l’élément actif.

##### <a name="read-mode"></a>Mode lecture

La propriété `to` renvoie un tableau contenant un objet `EmailAddressDetails` pour chaque destinataire répertorié sur la ligne **À** du message. Par défaut, la collection est limitée à 100 membres. Cependant, sous Windows et Mac, vous pouvez obtenir 500 membres au maximum.

```js
console.log(JSON.stringify(Office.context.mailbox.item.to));
```

##### <a name="compose-mode"></a>Mode composition

La propriété `to` renvoie un objet `Recipients` qui fournit des méthodes permettant d’obtenir ou de mettre à jour les destinataires figurant sur la ligne **À** du message. Par défaut, la collection est limitée à 100 membres. Toutefois, sous Windows et Mac, les limites suivantes s’appliquent.

- Obtenez 500 membres au maximum.
- Configurez un maximum de 100 membres par appel, jusqu’à 500 membres au total.

```js
Office.context.mailbox.item.to.setAsync( ['alice@contoso.com', 'bob@contoso.com'] );
Office.context.mailbox.item.to.addAsync( ['jason@contoso.com'] );
Office.context.mailbox.item.to.getAsync(callback);

function callback(asyncResult) {
  var arrayOfToRecipients = asyncResult.value;
}
```

##### <a name="type"></a>Type

*   Array.<[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)> | [Recipients](/javascript/api/outlook/office.recipients)

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

### <a name="methods"></a>Méthodes

#### <a name="addfileattachmentasyncuri-attachmentname-options-callback"></a>addFileAttachmentAsync(uri, attachmentName, [options], [callback])

Ajoute un fichier à un message ou un rendez-vous en pièce jointe.

La méthode `addFileAttachmentAsync` charge le fichier depuis l’URI spécifié et le joint à l’élément dans le formulaire de composition.

L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.

##### <a name="parameters"></a>Paramètres
|Nom|Type|Attributs|Description|
|---|---|---|---|
|`uri`|Chaîne||URI indiquant l’emplacement du fichier à joindre au message ou au rendez-vous. La longueur maximale est de 2 048 caractères.|
|`attachmentName`|String||Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;facultatif&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`options.isInline`|Boolean|&lt;optional&gt;|Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.<br/>En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`AttachmentSizeExceeded`|La pièce jointe dépasse la taille autorisée.|
|`FileTypeNotSupported`|La pièce jointe comporte une extension qui n’est pas autorisée.|
|`NumberOfAttachmentsExceeded`|Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="examples"></a>Exemples

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  var attachmentURL = "https://contoso.com/rtm/icon.png";
  Office.context.mailbox.item.addFileAttachmentAsync(attachmentURL, attachmentURL, options, callback);
}
```

L’exemple suivant montre comment ajouter un fichier image comme pièce jointe incorporée et comment la pièce jointe est affichée dans le corps du message.

```js
Office.context.mailbox.item.addFileAttachmentAsync(
  "http://i.imgur.com/WJXklif.png",
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addfileattachmentfrombase64asyncbase64file-attachmentname-options-callback"></a>addFileAttachmentFromBase64Async (base64File, attachmentName, [options], [callback])

Ajoute un fichier à partir du codage Base64 à un message ou un rendez-vous en pièce jointe.

La `addFileAttachmentFromBase64Async` méthode charge le fichier à partir du codage Base64 et l’associe à l’élément dans le formulaire de composition. Cette méthode renvoie l’identificateur de pièce jointe dans l’objet AsyncResult. Value.

L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`base64File`|Chaîne||Contenu encodé en base64 d’une image ou d’un fichier à ajouter à un message électronique ou à un événement.|
|`attachmentName`|String||Nom de la pièce jointe affiché lors de son chargement. La taille maximale est de 255 caractères.|
|`options`|Objet|&lt;facultatif&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;facultatif&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`options.isInline`|Boolean|&lt;optional&gt;|Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.<br/>En cas d’échec du téléchargement de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`AttachmentSizeExceeded`|La pièce jointe dépasse la taille autorisée.|
|`FileTypeNotSupported`|La pièce jointe comporte une extension qui n’est pas autorisée.|
|`NumberOfAttachmentsExceeded`|Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="examples"></a>Exemples

```js
Office.context.mailbox.item.addFileAttachmentFromBase64Async(
  base64String,
  "cute_bird.png",
  {
    isInline: true
  },
  function (asyncResult) {
    Office.context.mailbox.item.body.setAsync(
      "<p>Here's a cute bird!</p><img src='cid:cute_bird.png'>",
      {
        "coercionType": "html"
      },
      function (asyncResult) {
        // Do something here.
      });
  });
```

<br>

---
---

#### <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

ajoute un gestionnaire d’événements pour un événement pris en charge.

Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.

##### <a name="parameters"></a>Parameters

| Nom | Type | Attributs | Description |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Événement qui doit appeler le gestionnaire. |
| `handler` | Fonction || Fonction qui gère l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d’objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`. |
| `options` | Objet | &lt;facultatif&gt; | Littéral d’objet contenant une ou plusieurs des propriétés suivantes. |
| `options.asyncContext` | Objet | &lt;facultatif&gt; | Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel. |
| `callback` | fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture |

##### <a name="example"></a>Exemple

```js
function myHandlerFunction(eventarg) {
  if (eventarg.attachmentStatus === Office.MailboxEnums.AttachmentStatus.Added) {
    var attachment = eventarg.attachmentDetails;
    console.log("Event Fired and Attachment Added!");
    getAttachmentContentAsync(attachment.id, options, callback);
  }
}

Office.context.mailbox.item.addHandlerAsync(Office.EventType.AttachmentsChanged, myHandlerFunction, myCallback);
```

<br>

---
---

#### <a name="additemattachmentasyncitemid-attachmentname-options-callback"></a>addItemAttachmentAsync(itemId, attachmentName, [options], [callback])

Ajoute un élément Exchange, comme un message, en pièce jointe au message ou au rendez-vous.

La méthode `addItemAttachmentAsync` joint l’élément avec l’identificateur Exchange spécifié à l’élément du formulaire de composition. Si vous spécifiez une méthode de rappel, la méthode est appelée avec un paramètre, `asyncResult`, qui contient l’identificateur de pièce jointe ou un code indiquant toute erreur survenue lors de l’ajout de l’élément en tant que pièce jointe. Si nécessaire, vous pouvez utiliser le paramètre `options` pour transmettre des informations d’état à la méthode de rappel.

L’identificateur peut être utilisé avec la méthode [`removeAttachmentAsync`](#removeattachmentasyncattachmentid-options-callback) pour supprimer la pièce jointe dans la même session.

Si votre complément Office est exécuté dans Outlook sur le web, la méthode `addItemAttachmentAsync` peut joindre des éléments à des éléments autres que ceux que vous modifiez ; mais cette action n’est pas prise en charge et est déconseillée.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`itemId`|Chaîne||Identificateur Exchange de l’élément à joindre. La taille maximale est de 100 caractères.|
|`attachmentName`|String||Objet de l’élément à joindre. La longueur maximale est de 255 caractères.|
|`options`|Object|&lt;facultatif&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;facultatif&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>En cas de réussite, l’identificateur de pièce jointe est fourni dans la propriété `asyncResult.value`.<br/>En cas d’échec de l’ajout de la pièce jointe, l’objet `asyncResult` contient un objet `Error` indiquant une description de l’erreur.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`NumberOfAttachmentsExceeded`|Le message ou le rendez-vous comporte un trop grand nombre de pièces jointes.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="example"></a>Exemple

L’exemple suivant ajoute un élément Outlook existant en tant que pièce jointe avec le nom `My Attachment`.

```js
function callback(result) {
  if (result.error) {
    console.log(result.error);
  } else {
    console.log("Attachment added");
  }
}

function addAttachment() {
  // EWS ID of item to attach (shortened for readability).
  var itemId = "AAMkADI1...AAA=";

  // The values in asyncContext can be accessed in the callback.
  var options = { 'asyncContext': { var1: 1, var2: 2 } };

  Office.context.mailbox.item.addItemAttachmentAsync(itemId, "My Attachment", options, callback);
}
```

<br>

---
---

#### <a name="close"></a>close()

Ferme l’élément en cours qui est composé.

Le comportement de la méthode `close` dépend de l’état actuel de l’élément en cours de composition. Si l’élément comprend des modifications non enregistrées, le client invite l’utilisateur à enregistrer les modifications, à les ignorer ou à annuler l’action Fermer.

> [!NOTE]
> Dans Outlook sur le web, si l’élément est un rendez-vous et s’il a été précédemment enregistré à l’aide de `saveAsync`, l’utilisateur est invité à enregistrer, abandonner ou annuler même si aucune modification n’a été apportée depuis le dernier enregistrement de l’élément.

Dans le client de bureau Outlook, si le message est une réponse instantanée, la méthode `close` n’a aucun effet.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|Restreinte|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

<br>

---
---

#### <a name="displayreplyallformformdata-callback"></a>displayReplyAllForm(formData, [callback])

Affiche un formulaire de réponse qui inclut, soit l’expéditeur et tous les destinataires du message sélectionné, soit l’organisateur et tous les participants du rendez-vous sélectionné.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.

Si un des paramètres de chaîne dépasse la limite, `displayReplyAllForm` génère une exception.

Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`formData`|String &#124; Object||Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.<br/>**OU**<br/>Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :|
|`formData.htmlBody`|Chaîne|&lt;optional&gt;|Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;optional&gt;|Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.|
|`formData.attachments.type`|Chaîne||Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.|
|`formData.attachments.name`|Chaîne||Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.|
|`formData.attachments.url`|Chaîne||Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.|
|`formData.attachments.isInline`|Booléen||Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.|
|`formData.attachments.itemId`|String||Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.|
|`callback`|function|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="examples"></a>Exemples

Le code suivant transmet une chaîne à la fonction `displayReplyAllForm`.

```js
Office.context.mailbox.item.displayReplyAllForm('hello there');
Office.context.mailbox.item.displayReplyAllForm('<b>hello there</b>');
```

Réponse avec un corps vide.

```js
Office.context.mailbox.item.displayReplyAllForm({});
```

Réponse avec un corps.

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi'
});
```

Réponse avec un corps et la pièce jointe d’un fichier.

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Réponse avec un corps et la pièce jointe d’un élément.

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.

```js
Office.context.mailbox.item.displayReplyAllForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="displayreplyformformdata-callback"></a>displayReplyForm(formData, [callback])

Affiche un formulaire de réponse qui comprend uniquement l’expéditeur du message sélectionné ou l’organisateur du rendez-vous sélectionné.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

Dans Outlook sur le web, le formulaire de réponse s’affiche sous forme de formulaire isolé dans l’affichage à 3 colonnes et sous forme de formulaire contextuel dans l’affichage à 1 ou 2 colonnes.

Si un des paramètres de chaîne dépasse la limite, `displayReplyForm` génère une exception.

Lorsque des pièces jointes sont spécifiées dans le paramètre `formData.attachments`, Outlook et Outlook sur le web et clients bureau tentent de télécharger toutes les pièces jointes et de les joindre au formulaire de réponse. Si aucune pièce jointe n’est ajoutée, une erreur s’affiche dans l’interface utilisateur du formulaire. Si ce n’est pas possible, aucun message d’erreur n’est généré.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`formData`|String &#124; Object||Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.<br/>**OU**<br/>Objet qui contient les données du corps du message ou des pièces jointes et une fonction de rappel. L’objet est défini de la manière suivante :|
|`formData.htmlBody`|Chaîne|&lt;optional&gt;|Chaîne qui contient du texte et des éléments HTML et qui représente le corps du formulaire de réponse. La chaîne est limitée à 32 Ko.
|`formData.attachments`|Array.&lt;Object&gt;|&lt;optional&gt;|Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément.|
|`formData.attachments.type`|Chaîne||Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément.|
|`formData.attachments.name`|Chaîne||Chaîne qui contient le nom de la pièce jointe et comporte jusqu'à 255 caractères.|
|`formData.attachments.url`|Chaîne||Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier.|
|`formData.attachments.isInline`|Booléen||Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incorporée dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes.|
|`formData.attachments.itemId`|String||Utilisé uniquement si `type` est défini sur `item`. Il s’agit de l’ID de l’élément EWS de la pièce jointe. Il s’agit d’une chaîne comportant un maximum de 100 caractères.|
|`callback`|function|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="examples"></a>Exemples

Le code suivant transmet une chaîne à la fonction `displayReplyForm`.

```js
Office.context.mailbox.item.displayReplyForm('hello there');
Office.context.mailbox.item.displayReplyForm('<b>hello there</b>');
```

Réponse avec un corps vide.

```js
Office.context.mailbox.item.displayReplyForm({});
```

Réponse avec un corps.

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi'
});
```

Réponse avec un corps et la pièce jointe d’un fichier.

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    }
  ]
});
```

Réponse avec un corps et la pièce jointe d’un élément.

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ]
});
```

Réponse avec un corps, la pièce jointe d’un fichier, la pièce jointe d’un élément et un rappel.

```js
Office.context.mailbox.item.displayReplyForm(
{
  'htmlBody' : 'hi',
  'attachments' :
  [
    {
      'type' : Office.MailboxEnums.AttachmentType.File,
      'name' : 'squirrel.png',
      'url' : 'http://i.imgur.com/sRgTlGR.jpg'
    },
    {
      'type' : 'item',
      'name' : 'rand',
      'itemId' : Office.context.mailbox.item.itemId
    }
  ],
  'callback' : function(asyncResult)
  {
    console.log(asyncResult.value);
  }
});
```

<br>

---
---

#### <a name="getallinternetheadersasyncoptions-callback"></a>getAllInternetHeadersAsync ([options], [Rappel])

Obtient tous les en-têtes Internet pour le message sous forme de chaîne. Mode Lecture uniquement.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`options`|Objet|&lt;facultatif&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;facultatif&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult). En cas de réussite, les données des en-têtes Internet sont fournies dans la propriété asyncResult. Value sous forme de chaîne. Reportez-vous à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183) pour les informations de mise en forme de la valeur de chaîne renvoyée. En cas d’échec de l’appel, la propriété asyncResult. Error contient un code d’erreur correspondant à la raison de l’échec.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Les données des en-têtes Internet sous forme de chaîne formatée conformément à la [norme RFC 2183](https://tools.ietf.org/html/rfc2183).

Type : String

##### <a name="example"></a>Exemple

```js
// Get the internet headers related to the mail.
Office.context.mailbox.item.getAllInternetHeadersAsync(
  function(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
      console.log(asyncResult.value);
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getattachmentcontentasyncattachmentid-options-callback--attachmentcontentjavascriptapioutlookofficeattachmentcontent"></a>getAttachmentContentAsync (attachmentId, [options], [callback]) → [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

Obtient la pièce jointe spécifiée à partir d’un message ou d’un `AttachmentContent` rendez-vous et la renvoie en tant qu’objet.

La `getAttachmentContentAsync` méthode obtient la pièce jointe avec l’identificateur spécifié à partir de l’élément. Il est recommandé d’utiliser l’identificateur pour récupérer une pièce jointe dans la même session que l’attachmentIds a été récupérée avec l' `getAttachmentsAsync` appel ou `item.attachments` . Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`attachmentId`|Chaîne||Identificateur de la pièce jointe que vous souhaitez obtenir.|
|`options`|Objet|&lt;facultatif&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;facultatif&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="returns"></a>Renvoie :

Type : [AttachmentContent](/javascript/api/outlook/office.attachmentcontent)

##### <a name="example"></a>Exemple

```js
var item = Office.context.mailbox.item;
var listOfAttachments = [];
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

<br>

---
---

#### <a name="getattachmentsasyncoptions-callback--arrayattachmentdetailsjavascriptapioutlookofficeattachmentdetails"></a>getAttachmentsAsync ([options], [Rappel]) → Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

Obtient les pièces jointes de l’élément sous la forme d’un tableau. Mode composition uniquement.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`options`|Objet|&lt;facultatif&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [AsyncResult](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="returns"></a>Renvoie :

Type : Array. <[AttachmentDetails](/javascript/api/outlook/office.attachmentdetails)>

##### <a name="example"></a>Exemple

L’exemple suivant génère une chaîne HTML avec les détails de toutes les pièces jointes de l’élément actif.

```js
var item = Office.context.mailbox.item;
var outputString = "";
item.getAttachmentsAsync(callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      var attachment = result.value [i];
      outputString += "<BR>" + i + ". Name: ";
      outputString += attachment.name;
      outputString += "<BR>ID: " + attachment.id;
      outputString += "<BR>contentType: " + attachment.contentType;
      outputString += "<BR>size: " + attachment.size;
      outputString += "<BR>attachmentType: " + attachment.attachmentType;
      outputString += "<BR>isInline: " + attachment.isInline;
    }
  }
}
```

<br>

---
---

#### <a name="getentities--entitiesjavascriptapioutlookofficeentities"></a>getEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Permet d’obtenir les entités figurant dans le corps de l’élément sélectionné.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Type : [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Exemple

L’exemple suivant accède aux entités des contacts dans le corps de l’élément actif.

```js
var contacts = Office.context.mailbox.item.getEntities().contacts;
```

<br>

---
---

#### <a name="getentitiesbytypeentitytype--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getEntitiesByType(entityType) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Permet d’obtenir un tableau de toutes les entités du type spécifié trouvées dans le corps de l’élément sélectionné.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Description|
|---|---|---|
|`entityType`|[Office.MailboxEnums.EntityType](/javascript/api/outlook/office.mailboxenums.entitytype)|Une des valeurs d’énumération EntityType.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|Restreinte|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Si la valeur transmise à `entityType` n’est pas un membre valide de l’énumération `EntityType`, la méthode renvoie la valeur null. Si aucune entité du type spécifié n’est présente dans le corps de l’élément, la méthode renvoie un tableau vide. Sinon, le type des objets dans le tableau renvoyé dépend du type d’entité demandé dans le paramètre `entityType`.

Alors que le niveau d’autorisation minimal **Restricted** suffit pour utiliser cette méthode, certains types d’entité nécessitent le niveau **ReadItem** pour pouvoir y accéder, comme indiqué dans le tableau suivant.

|Valeur de `entityType`|Type des objets du tableau renvoyé|Niveau d’autorisation requis|
|---|---|---|
|`Address`|String|**Restricted**|
|`Contact`|Contact|**ReadItem**|
|`EmailAddress`|String|**ReadItem**|
|`MeetingSuggestion`|MeetingSuggestion|**ReadItem**|
|`PhoneNumber`|PhoneNumber|**Restricted**|
|`TaskSuggestion`|TaskSuggestion|**ReadItem**|
|`URL`|String|**Restricted**|

Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

##### <a name="example"></a>Exemple

L’exemple suivant montre comment accéder à un tableau de chaînes qui représente des adresses postales dans le corps de l’élément actif.

```js
// The initialize function is required for all apps.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, app-specific code can run.
    var item = Office.context.mailbox.item;
    // Get an array of strings that represent postal addresses in the current item's body.
    var addresses = item.getEntitiesByType(Office.MailboxEnums.EntityType.Address);
    // Continue processing the array of addresses.
  });
};
```

<br>

---
---

#### <a name="getfilteredentitiesbynamename--nullable-arraystringcontactjavascriptapioutlookofficecontactmeetingsuggestionjavascriptapioutlookofficemeetingsuggestionphonenumberjavascriptapioutlookofficephonenumbertasksuggestionjavascriptapioutlookofficetasksuggestion"></a>getFilteredEntitiesByName(name) → (nullable) {Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>}

Renvoie des entités reconnues dans l’élément sélectionné, qui transmettent le filtre nommé défini au fichier manifeste XML.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `getFilteredEntitiesByName` renvoie les entités qui correspondent à l’expression régulière définie dans l’élément de règle [ItemHasKnownEntity](/office/dev/add-ins/reference/manifest/rule#itemhasknownentity-rule) du fichier manifeste XML ayant la valeur de l’élément `FilterName` spécifié.

##### <a name="parameters"></a>Parameters

|Nom|Type|Description|
|---|---|---|
|`name`|String|Nom de l’élément de règle `ItemHasKnownEntity` qui définit le filtre à respecter.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Si aucun élément `ItemHasKnownEntity` dans le manifeste n’a une valeur d’élément `FilterName` qui correspond au paramètre `name`, la méthode renvoie `null`. Si le paramètre `name` correspond à un élément `ItemHasKnownEntity` dans le manifeste, mais qu’aucune entité dans l’élément actif ne correspond, la méthode renvoie un tableau vide.

Type : Array.<(String|[Contact](/javascript/api/outlook/office.contact)|[MeetingSuggestion](/javascript/api/outlook/office.meetingsuggestion)|[PhoneNumber](/javascript/api/outlook/office.phonenumber)|[TaskSuggestion](/javascript/api/outlook/office.tasksuggestion))>

<br>

---
---

#### <a name="getinitializationcontextasyncoptions-callback"></a>getInitializationContextAsync ([options], [Rappel])

Obtient les données d’initialisation transmises lorsque le complément est [activé par un message actionnable](/outlook/actionable-messages/invoke-add-in-from-actionable-message).

> [!NOTE]
> Cette méthode est uniquement prise en charge par Outlook 2016 ou une version ultérieure sur Windows (versions « démarrer en un clic » ultérieures à 16.0.8413.1000) et Outlook sur le Web pour Office 365.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>En cas de réussite, les données d’initialisation sont fournies `asyncResult.value` dans la propriété sous la forme d’une chaîne.<br/>S’il n’existe pas de contexte d’initialisation `asyncResult` , l’objet contient `Error` un objet dont `code` la propriété est `9020` définie sur `name` et sa propriété `GenericResponseError`est définie sur.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|Aperçu|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="example"></a>Exemple

```js
// Get the initialization context (if present).
Office.context.mailbox.item.getInitializationContextAsync(
  function(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      if (asyncResult.value != null && asyncResult.value.length > 0) {
        // The value is a string, parse to an object.
        var context = JSON.parse(asyncResult.value);
        // Do something with context.
      } else {
        // Empty context, treat as no context.
      }
    } else {
      if (asyncResult.error.code == 9020) {
        // GenericResponseError returned when there is no context.
        // Treat as no context.
      } else {
        // Handle the error.
      }
    }
  }
);
```

<br>

---
---

#### <a name="getitemidasyncoptions-callback"></a>getItemIdAsync ([options], rappel)

Obtient de manière asynchrone l’ID d’un élément enregistré. Mode composition uniquement.

Lorsqu’elle est appelée, cette méthode renvoie l’ID de l’élément par le biais de la méthode de rappel.

> [!NOTE]
> Si votre complément appelle `getItemIdAsync` sur un élément en mode composition (par exemple, pour obtenir un à utiliser avec `itemId` EWS ou l’API REST), sachez que lorsque Outlook est en mode mis en cache, l’élément peut prendre un certain temps avant la synchronisation de l’élément avec le serveur. Tant que l’élément n’est pas synchronisé `itemId` , le n’est pas reconnu et son utilisation renvoie une erreur.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`ItemNotSaved`|L’ID ne peut pas être récupéré tant que l’élément n’est pas enregistré.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="examples"></a>Exemples

```js
Office.context.mailbox.item.getItemIdAsync(
  function callback(result) {
    // Process the result.
  });
```

L’exemple suivant montre la structure du `result` paramètre transmis à la fonction de rappel. La `value` propriété contient l’ID de l’élément.

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="getregexmatches--object"></a>getRegExMatches() → {Object}

Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières définies dans le fichier manifeste XML.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `getRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.

Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.

<dl class="param-type">

<dt>Type</dt>

<dd>Object</dd>

</dl>

##### <a name="example"></a>Exemple

L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.

```js
var allMatches = Office.context.mailbox.item.getRegExMatches();
var fruits = allMatches.fruits;
var veggies = allMatches.veggies;
```

<br>

---
---

#### <a name="getregexmatchesbynamename--nullable-array-string-"></a>getRegExMatchesByName(name) → (nullable) {Array.< String >}

Renvoie des valeurs de chaîne dans l’élément sélectionné, qui correspondent aux expressions régulières nommées définies dans le fichier manifeste XML.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `getRegExMatchesByName` renvoie les chaînes qui correspondent à l’expression régulière définie dans l’élément de règle `ItemHasRegularExpressionMatch` du fichier manifeste XML ayant la valeur de l’élément `RegExName` spécifié.

Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Description|
|---|---|---|
|`name`|String|Nom de l’élément de règle `ItemHasRegularExpressionMatch` qui définit le filtre à respecter.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Un tableau qui contient les chaînes correspondant à l’expression régulière définie dans le fichier manifeste XML.

Type : Array.< String >

##### <a name="example"></a>Exemple

```js
var fruits = Office.context.mailbox.item.getRegExMatchesByName("fruits");
var veggies = Office.context.mailbox.item.getRegExMatchesByName("veggies");
```

<br>

---
---

#### <a name="getselecteddataasynccoerciontype-options-callback--string"></a>getSelectedDataAsync(coercionType, [options], callback) → {String}

Renvoie de manière asynchrone les données sélectionnées à partir de l’objet ou du corps d’un message.

Si aucune sélection n’est effectuée, mais que le curseur est placé dans le corps ou l’objet, la méthode renvoie la valeur null pour les données sélectionnées. Si un champ autre que le corps ou l’objet est sélectionné, la méthode renvoie l’erreur `InvalidSelection`.

> [!NOTE]
> Dans Outlook sur le Web, la méthode renvoie la chaîne « NULL » si aucun texte n’est sélectionné, mais que le curseur se trouve dans le corps. Pour vérifier cette situation, incluez un code similaire à celui-ci :
>
> `var selectedText = (asyncResult.value.endPosition === asyncResult.value.startPosition) ? "" : asyncResult.value.data;`

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`coercionType`|[Office.CoercionType](office.md#coerciontype-string)||Demande un format à attribuer aux données. S’il s’agit de texte, la méthode renvoie le texte brut en tant que chaîne, en retirant toutes les balises HTML présentes. S’il s’agit de langage HTML, la méthode renvoie le texte sélectionné, qu’il s’agisse de texte brut ou de langage HTML.|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Pour accéder aux données sélectionnées via la méthode de rappel, appelez la méthode `asyncResult.value.data`. Pour accéder à la propriété source dont la sélection est issue, appelez la méthode `asyncResult.value.sourceProperty`, qui correspond à `body` ou `subject`.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="returns"></a>Renvoie :

Les données sélectionnées en tant que chaîne dont le format est déterminé par `coercionType`.

Type : String

##### <a name="example"></a>Exemple

```js
// Get selected data.
Office.initialize = function () {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, {}, getCallback);
};

function getCallback(asyncResult) {
  var text = asyncResult.value.data;
  var prop = asyncResult.value.sourceProperty;

  Office.context.mailbox.item.setSelectedDataAsync('Setting ' + prop + ': ' + text, {}, setCallback);
}

function setCallback(asyncResult) {
  // Check for errors.
}
```

<br>

---
---

#### <a name="getselectedentities--entitiesjavascriptapioutlookofficeentities"></a>getSelectedEntities() → {[Entities](/javascript/api/outlook/office.entities)}

Obtient les entités figurant dans une correspondance en surbrillance qu’un utilisateur a sélectionné. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Type : [Entities](/javascript/api/outlook/office.entities)

##### <a name="example"></a>Exemple

L’exemple suivant accède aux entités d’adresses dans la correspondance en surbrillance sélectionnée par l’utilisateur.

```js
var contacts = Office.context.mailbox.item.getSelectedEntities().addresses;
```

<br>

---
---

#### <a name="getselectedregexmatches--object"></a>getSelectedRegExMatches() → {Object}

Renvoie des valeurs de chaîne dans une correspondance en surbrillance, qui correspondent aux expressions régulières définies dans le fichier manifeste XML. Les correspondances en surbrillance s’appliquent aux [compléments contextuels](/outlook/add-ins/contextual-outlook-add-ins).

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `getSelectedRegExMatches` renvoie les chaînes qui correspondent à l’expression régulière définie dans chaque élément de règle `ItemHasRegularExpressionMatch` ou `ItemHasKnownEntity` du fichier manifeste XML. Pour une règle `ItemHasRegularExpressionMatch`, une chaîne correspondante doit être présente dans la propriété de l’élément spécifié par cette règle. Le type simple `PropertyName` définit les propriétés prises en charge.

Par exemple, supposons qu’un manifeste de complément contienne l’élément `Rule` suivant :

```xml
<Rule xsi:type="RuleCollection" Mode="And">
  <Rule xsi:type="ItemIs" FormType="Read" ItemType="Message" />
  <Rule xsi:type="RuleCollection" Mode="Or">
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="fruits" RegExValue="apple|banana|coconut" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
    <Rule xsi:type="ItemHasRegularExpressionMatch" RegExName="veggies" RegExValue="tomato|onion|spinach|broccoli" PropertyName="BodyAsPlaintext" IgnoreCase="true" />
  </Rule>
</Rule>
```

L’objet renvoyé depuis `getRegExMatches` aurait deux propriétés : `fruits` et `veggies`.

```json
{
  'fruits': ['apple','banana','Banana','coconut'],
  'veggies': ['tomato','onion','spinach','broccoli']
}
```

Si vous spécifiez une règle `ItemHasRegularExpressionMatch` pour la propriété de corps d’un élément, l’expression régulière doit filtrer davantage le corps. Par ailleurs, elle ne doit pas tenter de renvoyer l’intégralité du corps de l’élément. L’utilisation d’une expression régulière telle que `.*` pour obtenir l’intégralité du corps d’un élément ne renvoie pas toujours les résultats attendus. Utilisez plutôt la méthode [`Body.getAsync`](/javascript/api/outlook/office.body#getasync-coerciontype--options--callback-) pour récupérer l’intégralité du corps de l’élément.

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.6|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Lecture|

##### <a name="returns"></a>Renvoie :

Un objet qui contient les tableaux des chaînes correspondant aux expressions régulières définies dans le fichier manifeste XML. Le nom de chaque tableau est égal à la valeur correspondante de l’attribut `RegExName` de la règle `ItemHasRegularExpressionMatch` correspondante ou de l’attribut `FilterName` de la règle `ItemHasKnownEntity` correspondante.

##### <a name="example"></a>Exemple

L’exemple suivant montre comment accéder au tableau de correspondances pour les éléments de règle d’expression régulière `fruits` et `veggies`, spécifiés dans le manifeste.

```js
var selectedMatches = Office.context.mailbox.item.getSelectedRegExMatches();
var fruits = selectedMatches.fruits;
var veggies = selectedMatches.veggies;
```

<br>

---
---

#### <a name="getsharedpropertiesasyncoptions-callback"></a>getSharedPropertiesAsync ([options], rappel)

Obtient les propriétés du rendez-vous ou du message sélectionné dans un dossier partagé, un calendrier ou une boîte aux lettres.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Les propriétés partagées sont fournies sous [`SharedProperties`](/javascript/api/outlook/office.sharedproperties) la forme d' `asyncResult.value` un objet dans la propriété. Cet objet peut être utilisé pour obtenir les propriétés partagées de l’élément.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.8|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
Office.context.mailbox.item.getSharedPropertiesAsync(callback);

function callback (asyncResult) {
  var context = asyncResult.context;
  var sharedProperties = asyncResult.value;
}
```

<br>

---
---

#### <a name="loadcustompropertiesasynccallback-usercontext"></a>loadCustomPropertiesAsync(callback, [userContext])

Charge de manière asynchrone les propriétés personnalisées de ce complément vers l’élément sélectionné.

Les propriétés personnalisées sont stockées sous la forme de paires clé/valeur qui s’appliquent à une application ou un élément. Cette méthode renvoie un objet `CustomProperties` dans le rappel, qui fournit des méthodes pour accéder aux propriétés personnalisées propres à l’élément et au complément actifs. Les propriétés personnalisées ne sont pas chiffrées dans l’élément, par conséquent elles ne doivent pas servir d’espace de stockage sécurisé.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`callback`|function||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Les propriétés personnalisées sont fournies sous la forme d’un objet [`CustomProperties`](/javascript/api/outlook/office.customproperties) dans la propriété `asyncResult.value`. Cet objet peut être utilisé pour obtenir, définir et supprimer des propriétés personnalisées à partir de l’élément et réenregistrer les modifications apportées au jeu de propriétés personnalisées sur le serveur.|
|`userContext`|Objet|&lt;optional&gt;|Les développeurs peuvent fournir un objet auquel ils souhaitent accéder dans la fonction de rappel. Cet objet est accessible via la propriété `asyncResult.asyncContext` dans la fonction de rappel.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Rédaction ou lecture|

##### <a name="example"></a>Exemple

L’exemple de code suivant montre comment utiliser la méthode `loadCustomPropertiesAsync` pour charger de manière asynchrone des propriétés personnalisées spécifiques vers l’élément actif. L’exemple montre également comment utiliser la méthode `CustomProperties.saveAsync` pour réenregistrer ces propriétés sur le serveur. Une fois le chargement des propriétés personnalisées terminé, l’exemple de code utilise la méthode `CustomProperties.get` pour lire la propriété personnalisée `myProp`, utilise la méthode `CustomProperties.set` pour écrire la propriété personnalisée `otherProp`, puis appelle enfin la méthode `saveAsync` pour enregistrer les propriétés personnalisées.

```js
// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Checks for the DOM to load using the jQuery ready function.
  $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    var item = Office.context.mailbox.item;
    item.loadCustomPropertiesAsync(customPropsCallback);
  });
};

function customPropsCallback(asyncResult) {
  var customProps = asyncResult.value;
  var myProp = customProps.get("myProp");

  customProps.set("otherProp", "value");
  customProps.saveAsync(saveCallback);
}

function saveCallback(asyncResult) {
}
```

<br>

---
---

#### <a name="removeattachmentasyncattachmentid-options-callback"></a>removeAttachmentAsync(attachmentId, [options], [callback])

Supprime une pièce jointe d’un message ou d’un rendez-vous.

La méthode `removeAttachmentAsync` supprime la pièce jointe avec l’identificateur spécifié de l’élément. Nous vous recommandons vivement de supprimer une pièce jointe à l’aide de son identificateur uniquement si la même application de messagerie a ajouté cette pièce jointe au cours de la même session. Dans Outlook sur le web et sur les appareils mobiles, l’identificateur de pièce jointe n’est valable que dans la même session. Une session est terminée lorsque l’utilisateur ferme l’application, ou si l’utilisateur commence à composer un formulaire inséré, puis détoure ensuite le formulaire pour continuer dans une fenêtre distincte.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`attachmentId`|String||Identificateur de la pièce jointe à supprimer.|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction|&lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). <br/>En cas d’échec de la suppression de la pièce jointe, la propriété `asyncResult.error` contient un code d’erreur et la raison de l’échec.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`InvalidAttachmentId`|L’identificateur de la pièce jointe n’existe pas.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.1|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="example"></a>Exemple

Le code suivant supprime une pièce jointe dont l’identificateur est « 0 ».

```js
Office.context.mailbox.item.removeAttachmentAsync(
  '0',
  { asyncContext : null },
  function (asyncResult)
  {
    console.log(asyncResult.status);
  }
);
```

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a>removeHandlerAsync(eventType, [options], [callback])

Supprime les gestionnaires d’événements pour un type d’événement pris en charge.

Actuellement, les types d’événement `Office.EventType.AttachmentsChanged`pris `Office.EventType.AppointmentTimeChanged`en `Office.EventType.EnhancedLocationsChanged`charge `Office.EventType.RecipientsChanged`sont, `Office.EventType.RecurrenceChanged`,, et.

##### <a name="parameters"></a>Parameters

| Nom | Type | Attributs | Description |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || Événement qui doit révoquer le gestionnaire. |
| `options` | Objet | &lt;optional&gt; | Littéral d’objet contenant une ou plusieurs des propriétés suivantes. |
| `options.asyncContext` | Objet | &lt;optional&gt; | Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel. |
| `callback` | fonction| &lt;optional&gt;|Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.7 |
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture |

<br>

---
---

#### <a name="saveasyncoptions-callback"></a>saveAsync([options], callback)

Enregistre un élément de manière asynchrone.

Lorsqu’elle est appelée, cette méthode enregistre le message en cours en tant que brouillon et renvoie l’ID de l’élément via la méthode de rappel. Dans Outlook sur le web ou Outlook en mode en ligne, l’élément est enregistré sur le serveur. Dans Outlook en mode mis en cache, l’élément est enregistré dans le cache local.

> [!NOTE]
> Si votre complément appelle `saveAsync` sur un élément en mode composition afin d’obtenir un `itemId` à utiliser avec EWS ou l’API REST, sachez que lorsqu’Outlook est en mode mis en cache, la synchronisation de l’élément sur le serveur peut prendre un certain temps. Avant que l’élément ne soit synchronisé, l’utilisation de l’élément `itemId` renvoie une erreur.

Dans la mesure où les rendez-vous n’ont pas d’état brouillon, si `saveAsync` est appelé sur un rendez-vous en mode composition, l’élément est enregistré sous la forme d’un rendez-vous normal sur le calendrier de l’utilisateur. Pour les nouveaux rendez-vous qui n’ont pas été enregistrés, aucune invitation ne sera envoyée. L’enregistrement d’un rendez-vous existant envoie une mise à jour aux participants ajoutés ou supprimés.

> [!NOTE]
> Les clients suivants ont un comportement différent avec `saveAsync` sur les rendez-vous en mode composition :
>
> - Outlook pour Mac ne prend pas en charge l’enregistrement d’une réunion. La méthode `saveAsync` échoue lorsqu’elle est appelée à partir d’une réunion en mode composition. Pour contourner ce problème, voir [Impossible d’enregistrer une réunion en tant que brouillon dans Outlook pour Mac à l’aide des API de JS Office](https://support.microsoft.com/help/4505745).
> - Outlook sur le web envoie toujours une invitation ou une mise à jour quand `saveAsync` est appelée sur un rendez-vous en mode composition.

##### <a name="parameters"></a>Parameters

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`options`|Object|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`callback`|fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>En cas de réussite, l’identificateur de l’élément est fourni dans la propriété `asyncResult.value`.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.3|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="examples"></a>範例

```js
Office.context.mailbox.item.saveAsync(
  function callback(result) {
    // Process the result.
  });
```

Voici un exemple du paramètre `result` transmis à la fonction de rappel. La propriété `value` contient l’ID de l’élément.

```json
{
  "value":"AAMkADI5...AAA=",
  "status":"succeeded"
}
```

<br>

---
---

#### <a name="setselecteddataasyncdata-options-callback"></a>setSelectedDataAsync(data, [options], callback)

Insère les données dans le corps ou l’objet d’un message de manière asynchrone.

La méthode `setSelectedDataAsync` insère la chaîne spécifiée à l’emplacement du curseur dans le corps ou l’objet de l’élément. En revanche, si du texte est sélectionné dans l’éditeur, il remplace le texte sélectionné. Si le curseur ne se trouve pas dans le champ du corps ou de l’objet, une erreur est renvoyée. Après l’insertion, le curseur est placé à la fin du contenu inséré.

##### <a name="parameters"></a>Paramètres

|Nom|Type|Attributs|Description|
|---|---|---|---|
|`data`|String||Données à insérer. Les données ne doivent pas dépasser 1 000 000 caractères. Si elles contiennent plus de 1 000 000 caractères, une exception `ArgumentOutOfRange` est générée.|
|`options`|Objet|&lt;optional&gt;|Littéral d’objet contenant une ou plusieurs des propriétés suivantes.|
|`options.asyncContext`|Objet|&lt;optional&gt;|Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel.|
|`options.coercionType`|[Office.CoercionType](office.md#coerciontype-string)|&lt;optional&gt;|Si `text`, le style existant est appliqué dans Outlook sur le web et Outlook client bureau. Si le champ est un éditeur HTML, seules les données de texte sont insérées, même si les données sont au format HTML.<br/><br/>Avec `html` et si le champ prend en charge le langage HTML (contrairement à l’objet), le style existant est appliqué dans Outlook sur le web et le style par défaut dans Outlook bureau. Si le champ est au format texte, une erreur `InvalidDataFormat` est renvoyée.<br/><br/>Si la propriété `coercionType` n’est pas définie, le résultat dépend du champ : si le champ est au format HTML, le langage HTML est utilisé ; si le champ est au format texte, le texte brut est utilisé.|
|`callback`|fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Configuration requise

|Conditions requises|Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)|1.2|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)|ReadWriteItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)|Composition|

##### <a name="example"></a>Exemple

```js
Office.context.mailbox.item.setSelectedDataAsync("Hello World!");
Office.context.mailbox.item.setSelectedDataAsync("<b>Hello World!</b>", { coercionType : "html" });
```
