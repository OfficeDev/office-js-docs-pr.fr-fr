---
title: Office. Context. Mailbox-ensemble de conditions requises 1,2
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 542e8c9899c2d4a3c5b4546c3d5a73ba0d3c3a7e
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/22/2019
ms.locfileid: "37626998"
---
# <a name="mailbox"></a>boîte aux lettres

### <a name="officeofficemdcontextofficecontextmdmailbox"></a>[Office](Office.md)[.context](Office.context.md).mailbox

Permet d’accéder au modèle d’objet de complément Outlook pour Microsoft Outlook.

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| Restreinte|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [ewsUrl](#ewsurl-string) | Membre |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Méthode |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Méthode |
| [displayAppointmentForm](#displayappointmentformitemid) | Méthode |
| [displayMessageForm](#displaymessageformitemid) | Méthode |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Méthode |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Méthode |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Méthode |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Méthode |

### <a name="namespaces"></a>Espaces de noms

[diagnostics](Office.context.mailbox.diagnostics.md) : Fournit des informations de diagnostic à un complément Outlook.

[item](Office.context.mailbox.item.md) : Fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.

[userProfile](Office.context.mailbox.userProfile.md) : Fournit des informations sur l’utilisateur dans un complément Outlook.

### <a name="members"></a>Members

#### <a name="ewsurl-string"></a>ewsUrl: String

Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.

> [!NOTE]
> Ce membre n’est pas pris en charge dans Outlook pour iOS ou Outlook pour Android.

La valeur `ewsUrl` peut être utilisée par un service distant pour émettre des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).

##### <a name="type"></a>Type

*   String

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Lecture|

### <a name="methods"></a>Méthodes

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}

Obtient un dictionnaire contenant les informations d’heure dans l’heure locale du client.

Une application de messagerie pour Outlook ou Outlook sur le web peut utiliser des fuseaux horaires différents pour les dates et heures. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure afin que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.

Si l’application de messagerie est en cours d’exécution dans Outlook sur ordinateur, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie est en cours d’exécution dans Outlook sur le web, la méthode `convertToLocalClientTime` renvoie un objet de dictionnaire dont les valeurs sont définies pour le fuseau horaire spécifié dans le CAE.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Description|
|---|---|---|
|`timeValue`| Date|Objet Date|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="returns"></a>Renvoie :

Type : [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.

La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs appropriées pour la date et l’heure locales.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Description|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|Valeur de l’heure locale à convertir.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="returns"></a>Renvoie :

Objet Date avec l’heure exprimée au format UTC.

Type : Date

##### <a name="example"></a>Exemple

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Affiche un rendez-vous de calendrier existant.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `displayAppointmentForm` ouvre un rendez-vous du calendrier existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.

Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. En effet, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (notamment l’ID d’élément) des instances d’une série périodique.

Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.

Si l’identificateur de l’élément spécifié n’identifie aucun rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client. Par ailleurs, aucun message d’erreur n’est retourné.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Description|
|---|---|---|
|`itemId`| String|Identificateur des services web Exchange pour un rendez-vous du calendrier existant.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Affiche un message existant.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre du Bureau ou dans une boîte de dialogue sur les appareils mobiles.

Dans Outlook sur le web, cette méthode ouvre le formulaire spécifié uniquement si le corps du formulaire comprend 32 Ko de caractères maximum.

Si l’identificateur de l’élément spécifié n’identifie aucun message existant, aucun message ne s’affiche sur l’ordinateur client. Par ailleurs, aucun message d’erreur n’est retourné.

N’utilisez pas la méthode `displayMessageForm` ayant une valeur `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire afin de créer un nouveau rendez-vous.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Description|
|---|---|---|
|`itemId`| String|Identificateur des services web Exchange pour un message existant.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Affiche un formulaire permettant de créer un rendez-vous du calendrier.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont spécifiés, les champs du formulaire de rendez-vous sont remplis automatiquement avec le contenu des paramètres.

Dans Outlook sur le web et appareils mobiles, cette méthode affiche toujours un formulaire contenant un champ Participants. Si vous ne spécifiez pas de participants comme arguments d’entrée, la méthode affiche un formulaire contenant le bouton **Enregistrer**. Si vous avez spécifié des participants, le formulaire inclut ces derniers, en plus du bouton **Envoyer**.

Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion comportant un bouton **Envoyer**. Si vous ne spécifiez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.

Si l’un des paramètres dépasse les limites définies en matière de taille ou si un nom de paramètre inconnu est spécifié, une exception est levée.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Description|
|---|---|---|
| `parameters` | Object | Dictionnaire de paramètres décrivant le nouveau rendez-vous. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt; | Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis du rendez-vous. Le tableau est limité à 100 entrées maximum. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt; | Tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à 100 entrées maximum. |
| `parameters.start` | Date | Objet `Date` spécifiant la date et l’heure de début du rendez-vous. |
| `parameters.end` | Date | Objet `Date` spécifiant la date et l’heure de fin du rendez-vous. |
| `parameters.location` | String | Chaîne contenant l’emplacement du rendez-vous. La chaîne est limitée à 255 caractères maximum. |
| `parameters.resources` | Array.&lt;String&gt; | Tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à 100 entrées maximum. |
| `parameters.subject` | String | Chaîne contenant l’objet du rendez-vous. La chaîne est limitée à 255 caractères maximum. |
| `parameters.body` | String | Corps du rendez-vous. La taille du corps du message est limitée à 32 Ko. |

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Lecture|

##### <a name="example"></a>Exemple

```js
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un serveur Exchange.

La méthode `getCallbackTokenAsync` émet un appel asynchrone pour obtenir un jeton opaque à partir du serveur Exchange qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.

Vous pouvez transmettre le jeton et soit un identificateur de pièce jointe, soit un identificateur d’élément à un système tiers. Le système tiers utilise le jeton comme jeton d’autorisation du support pour appeler l’opération [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) de services Web Exchange (EWS) ou de [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) pour renvoyer une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](/outlook/add-ins/get-attachments-of-an-outlook-item).

L’appel `getCallbackTokenAsync` de la méthode nécessite un niveau d’autorisation minimum de **ReadItem**.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| function||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.<br><br>En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.|
|`userContext`| Objet| &lt;optional&gt;|Données d’état transmises à la méthode asynchrone.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`HTTPRequestFailure`|La demande a échoué. Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.|
|`InternalServerError`|Le serveur Exchange a renvoyé une erreur. Pour plus d’informations, veuillez consulter l’objet de diagnostics.|
|`NetworkError`|L’utilisateur n’est plus connecté au réseau. Veuillez vérifier la connexion réseau et réessayer.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Lecture|

##### <a name="example"></a>Exemple

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Obtient un jeton qui identifie l’utilisateur et le complément Office.

La méthode `getUserIdentityTokenAsync` renvoie un jeton qui vous permet d’identifier et d’[authentifier le complément et l’utilisateur à l’aide d’un système tiers](/outlook/add-ins/authentication).

##### <a name="parameters"></a>Paramètres

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| function||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.<br><br>En cas d’erreur, les propriétés `asyncResult.error` et `asyncResult.diagnostics` peuvent fournir des informations supplémentaires.|
|`userContext`| Objet| &lt;optional&gt;|Données d’état transmises à la méthode asynchrone.|

##### <a name="errors"></a>Erreurs

|Code d'erreur|Description|
|------------|-------------|
|`HTTPRequestFailure`|La demande a échoué. Veuillez rechercher le code d’erreur HTTP dans l’objet de diagnostics.|
|`InternalServerError`|Le serveur Exchange a renvoyé une erreur. Pour plus d’informations, veuillez consulter l’objet de diagnostics.|
|`NetworkError`|L’utilisateur n’est plus connecté au réseau. Veuillez vérifier la connexion réseau et réessayer.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Envoie une demande asynchrone à un des services web Exchange (EWS) sur le serveur Exchange qui héberge la boîte aux lettres de l’utilisateur.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans les cas suivants :
> - Dans Outlook sur iOS ou Android
> - quand le complément est chargé dans une boîte aux lettres Gmail.
> 
> Dans ces cas de figure, les compléments doivent [utiliser les API REST](/outlook/add-ins/use-rest-api) pour accéder à la boîte aux lettres de l’utilisateur.

La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément. Pour obtenir la liste des opérations EWS prises en charge, reportez-vous à l’article [Appeler des services web à partir d’un complément Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).

Vous ne pouvez pas demander des éléments associés à un dossier avec la méthode `makeEwsRequestAsync`.

La demande XML doit spécifier l’encodage UTF-8.

```xml
<?xml version="1.0" encoding="utf-8"?>
```

Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et des opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, consultez la page relative aux[ autorisations du complément de messagerie pour accéder à la boîte aux lettres de l’utilisateur](/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> L’administrateur serveur doit définir `OAuthAuthentication` sur true dans le répertoire EWS du serveur d’accès client pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.

##### <a name="version-differences"></a>Différences entre les versions

Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie exécutées dans des versions d’Outlook inférieures à 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

Lorsque votre application de messagerie s’exécute dans Outlook sur le web, vous n’avez pas à définir la valeur d’encodage. Pour déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web, utilisez la propriété mailbox.diagnostics.hostName. Pour déterminer la version d’Outlook qui est exécutée, utilisez la propriété mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Paramètres

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||Demande EWS.|
|`callback`| function||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le résultat XML de l’appel EWS est fourni sous forme de chaîne dans la propriété `asyncResult.value`. Si la taille du résultat est supérieure à 1 Mo, un message d’erreur est renvoyé.|
|`userContext`| Objet| &lt;optional&gt;|Données d’état transmises à la méthode asynchrone.|

##### <a name="requirements"></a>Configuration requise

|Conditions requises| Valeur|
|---|---|
|[Version de l’ensemble minimal de conditions de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Mode Outlook applicable](/outlook/add-ins/#extension-points)| Rédaction ou lecture|

##### <a name="example"></a>Exemple

L’exemple suivant appelle la méthode `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.

```js
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```
