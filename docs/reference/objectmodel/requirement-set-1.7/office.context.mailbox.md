
# <a name="mailbox"></a>mailbox

### [Office](Office.md)[.context](Office.context.md).mailbox

Donne accès au modèle d’objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restreint|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="members-and-methods"></a>Membres et méthodes

| Membre | Type |
|--------|------|
| [ewsUrl](#ewsurl-string) | Membre |
| [restUrl](#resturl-string) | Membre |
| [addHandlerAsync](#addhandlerasynceventtype-handler-options-callback) | Méthode |
| [convertToEwsId](#converttoewsiditemid-restversion--string) | Méthode |
| [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | Méthode |
| [convertToRestId](#converttorestiditemid-restversion--string) | Méthode |
| [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Méthode |
| [displayAppointmentForm](#displayappointmentformitemid) | Méthode |
| [displayMessageForm](#displaymessageformitemid) | Méthode |
| [displayNewAppointmentForm](#displaynewappointmentformparameters) | Méthode |
| [displayNewMessageForm](#displaynewmessageformparameters) | Méthode |
| [getCallbackTokenAsync](#getcallbacktokenasyncoptions-callback) | Méthode |
| [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Méthode |
| [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Méthode |
| [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Méthode |

### <a name="namespaces"></a>Espaces de noms

[diagnostics](Office.context.mailbox.diagnostics.md) : fournit des informations de diagnostic à un complément Outlook.

[item](Office.context.mailbox.item.md) : fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.

[userProfile](Office.context.mailbox.userProfile.md) : fournit des informations sur l’utilisateur dans un complément Outlook.

### <a name="members"></a>Membres

#### <a name="ewsurl-string"></a>ewsUrl : chaîne

Obtient l’URL du point de terminaison des services Web Exchange (EWS) pour ce compte de messagerie. Mode lecture uniquement.

> [!NOTE]
> Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.

La valeur `ewsUrl` peut être utilisée par un service distant pour effectuer des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir les pièces jointes de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `ewsUrl` en mode lecture.

En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `ewsUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.

##### <a name="type"></a>Type :

*   String

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

#### <a name="resturl-string"></a>restUrl :String

Obtient l’URL du point de terminaison REST de ce compte de courrier.

La valeur `restUrl` peut être utilisée pour que l’[API REST](https://docs.microsoft.com/outlook/rest/) appelle la boîte aux lettres de l’utilisateur.

L’autorisation **ReadItem** doit être spécifiée dans le manifeste de votre application pour appeler le membre `restUrl` en mode lecture.

En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) avant de pouvoir utiliser le membre `restUrl`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.

##### <a name="type"></a>Type :

*   String

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

### <a name="methods"></a>Méthodes

####  <a name="addhandlerasynceventtype-handler-options-callback"></a>addHandlerAsync(eventType, handler, [options], [callback])

Ajoute un gestionnaire d’événements pour un événement pris en charge.

Pour le moment, les types d’événements pris en charge sont `Office.EventType.ItemChanged` et `Office.EventType.OfficeThemeChanged`.

##### <a name="parameters"></a>Paramètres :

| Name | Type | Attributs | Description |
|---|---|---|---|
| `eventType` | [Office.EventType](office.md#eventtype-string) || L’événement qui doit invoquer le gestionnaire. |
| `handler` | Fonction || La fonction permettant de gérer l’événement. Cette fonction doit accepter un seul paramètre, qui est un littéral d'objet. La propriété `type` sur le paramètre correspond au paramètre `eventType` transmis à `addHandlerAsync`. |
| `options` | Objet | &lt;facultatif&gt; | Littéral d’objet contenant une ou plusieurs des propriétés suivantes. |
| `options.asyncContext` | Objet | &lt;facultatif&gt; | Les développeurs peuvent indiquer un objet auquel ils souhaitent accéder dans la méthode de rappel. |
| `callback` | fonction| &lt;facultatif&gt;|Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem |
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a>convertToEwsId(itemId, restVersion) → {String}

Convertit un ID d’élément mis en forme pour REST au format EWS.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

Les ID d’éléments extraits via une API REST (telle que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)) utilisent un format différent de celui employé par les services Web Exchange (EWS). La méthode `convertToEwsId` convertit un ID mis en forme pour REST au format approprié pour EWS.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Description|
|---|---|---|
|`itemId`| String|Un ID d’élément mis en forme pour les API REST Outlook|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|Valeur indiquant la version de l’API REST Outlook utilisée pour récupérer l’ID d’élément.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restreint|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="returns"></a>Retourne :

Type : String

##### <a name="example"></a>Exemple

```
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}

Obtient un dictionnaire contenant des informations d’heure dans l’heure locale du client.

Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur ; Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure de telle sorte que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire attendu par l’utilisateur.

Si l’application de messagerie s'exécute dans Outlook, la méthode `convertToLocalClientTime` retournera un objet de dictionnaire dont les valeurs seront définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie s’exécute dans Outlook Web App, la méthode `convertToLocalClientTime` retournera objet dictionnaire dont les valeurs seront définies pour le fuseau horaire spécifié dans le CAE.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Description|
|---|---|---|
|`timeValue`| Date|Un objet Date|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="returns"></a>Retourne :

Type : [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)

####  <a name="converttorestiditemid-restversion--string"></a>convertToRestId(itemId, restVersion) → {String}

Convertit un ID d’élément mis en forme pour EWS au format REST.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

Les ID d’éléments récupérés via EWS ou via la propriété `itemId` utilisent un format différent de celui employé par les API REST (telles que l’[API Courrier Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) ou [Microsoft Graph](http://graph.microsoft.io/)). La méthode `convertToRestId` convertit un ID mis en forme pour EWS dans un format adapté à REST.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Description|
|---|---|---|
|`itemId`| String|Un ID d’élément mis en forme pour les services Web Exchange (EWS)|
|`restVersion`| [Office.MailboxEnums.RestVersion](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|Valeur indiquant la version de l’API REST Outlook avec laquelle l’ID converti sera utilisé.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restreint|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="returns"></a>Retourne :

Type : String

##### <a name="example"></a>Exemple

```
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.

La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs correctes pour la date et l’heure locales.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Description|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)|Valeur en heure locale à convertir.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="returns"></a>Retourne :

Un objet Date avec l’heure exprimée en UTC.

<dl class="param-type">

<dt>Type</dt>

<dd>Date</dd>

</dl>

####  <a name="displayappointmentformitemid"></a>displayAppointmentForm(itemId)

Affiche un rendez-vous de calendrier existant.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `displayAppointmentForm` ouvre un rendez-vous de calendrier existant dans une nouvelle fenêtre sur le bureau ou dans une boîte de dialogue sur les appareils mobiles.

Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. Cela est dû au fait que dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.

Dans Outlook Web App, cette méthode ouvre le formulaire spécifié seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.

Si l’identificateur de l’élément indiqué n’identifie pas un rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client, et aucun message d’erreur ne sera retourné.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Description|
|---|---|---|
|`itemId`| String|L'identificateur des services web Exchange pour un rendez-vous de calendrier existant.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a>displayMessageForm(itemId)

Affiche un message existant.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `displayMessageForm` ouvre un message existant dans une nouvelle fenêtre sur le bureau, ou dans une boîte de dialogue sur les appareils mobiles.

Dans Outlook Web App, cette méthode ouvre le formulaire indiqué seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.

Si l’identificateur de l’élément indiqué n’identifie pas un message existant, aucun message ne sera affiché sur l’ordinateur client, et aucun message d’erreur ne sera retourné.

N’utilisez pas la méthode `displayMessageForm` avec un `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire pour créer un nouveau rendez-vous.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Description|
|---|---|---|
|`itemId`| String|Identificateur des services web Exchange pour un message existant.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a>displayNewAppointmentForm(parameters)

Affiche un formulaire pour créer un rendez-vous de calendrier.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans Outlook pour iOS ou Outlook pour Android.

La méthode `displayNewAppointmentForm` ouvre un formulaire qui permet à l’utilisateur de créer un rendez-vous ou une réunion. Si des paramètres sont fournis, les champs du formulaire de rendez-vous sont automatiquement remplis avec le contenu des paramètres.

Dans Outlook Web App et OWA for Devices, cette méthode affiche toujours un formulaire avec un champ participants. Si vous n'indiquez aucun participant dans les arguments d’entrée, la méthode affiche un formulaire avec un bouton **Enregistrer**. Si vous avez indiqué des participants, le formulaire incluera les participants et un bouton **Envoyer**.

Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans le paramètre `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion avec un bouton **Envoyer**. Si vous ne n'indiquez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.

Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.

##### <a name="parameters"></a>Paramètres :

> [!NOTE]
> Tous les paramètres sont facultatifs.

|Name| Type| Description|
|---|---|---|
| `parameters` | Objet | Un dictionnaire de paramètres décrivant le nouveau rendez-vous. |
| `parameters.requiredAttendees` | Tableau.&lt;Chaîne&gt; | Tableau.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt; | Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées. |
| `parameters.optionalAttendees` | Tableau.&lt;Chaîne&gt; | Tableau.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt; | Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à un maximum de 100 entrées. |
| `parameters.start` | Date | Un objet `Date` indiquant la date et l’heure du début du rendez-vous. |
| `parameters.end` | Date | Un objet `Date` indiquant la date et l’heure de la fin du rendez-vous. |
| `parameters.location` | String | Un chaîne contenant le lieu du rendez-vous. La chaîne est limitée à un maximum de 255 caractères. |
| `parameters.resources` | Array.&lt;String&gt; | Un tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées. |
| `parameters.subject` | String | Une chaîne contenant l’objet du rendez-vous. La chaîne est limitée àun maximum de 255 caractères. |
| `parameters.body` | String | Le corps du rendez-vous. La contenu du corps est limitée à une taille maximale de 32 Ko. |

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lecture|

##### <a name="example"></a>Exemple

```
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

#### <a name="displaynewmessageformparameters"></a>displayNewMessageForm(parameters)

Affiche un formulaire permettant de créer un nouveau message.

La méthode `displayNewMessageForm` ouvre un formulaire qui permet à l’utilisateur de créer un nouveau message. Si des paramètres sont spécifiés, les champs du formulaire de message sont remplis automatiquement avec le contenu des paramètres.

Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.

##### <a name="parameters"></a>Paramètres :

> [!NOTE]
> Tous les paramètres sont facultatifs.

|Name| Type| Description|
|---|---|---|
| `parameters` | Objet | Dictionnaire de paramètres décrivant le nouveau message. |
| `parameters.toRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt; | Tableau de chaînes contenant les adresses e-mail ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne À. Le tableau est limité à 100 entrées maximum. |
| `parameters.ccRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt; | Tableau de chaînes contenant les adresses e-mail ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cc. Le tableau est limité à 100 entrées maximum. |
| `parameters.bccRecipients` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt; | Tableau de chaînes contenant les adresses e-mail ou tableau contenant un objet `EmailAddressDetails` pour chacun des destinataires de la ligne Cci. Le tableau est limité à 100 entrées maximum. |
| `parameters.subject` | String | Chaîne contenant l’objet du message. La chaîne est limitée à 255 caractères maximum. |
| `parameters.htmlBody` | String | Corps HTML du message. Le contenu du corps du message est limitée à une taille maximum de 32 Ko. |
| `parameters.attachments` | Array.&lt;Object&gt; | Tableau d’objets JSON qui sont des pièces jointes de fichier ou d’élément. |
| `parameters.attachments.type` | String | Indique le type de pièce jointe. Doit être `file` pour une pièce jointe de fichier ou `item` pour une pièce jointe d’élément. |
| `parameters.attachments.name` | String | Chaîne qui contient le nom de la pièce jointe, d’une longueur maximale de 255 caractères.|
| `parameters.attachments.url` | String | Utilisé uniquement si `type` est défini sur `file`. Il s’agit de l’URI de l’emplacement du fichier. |
| `parameters.attachments.isInline` | Boolean | Utilisé uniquement si `type` est défini sur `file`. Si elle est définie sur `true`, cette valeur indique que la pièce jointe est incluse dans le corps du message et qu’elle ne doit pas figurer dans la liste des pièces jointes. |
| `parameters.attachments.itemId` | String | Utilisé uniquement si `type` est défini sur `item`. L’id d’élément EWS du courrier électronique existant à joindre au nouveau message. Il s’agit d’une chaîne comportant un maximum de 100 caractères. |


##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.6 |
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lecture|

##### <a name="example"></a>Exemple

```
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

#### <a name="getcallbacktokenasyncoptions-callback"></a>getCallbackTokenAsync([options], callback)

Obtient une chaîne contenant un jeton utilisé pour appeler les API REST ou les services Web Exchange.

La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir de l’Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.

> [!NOTE]
> Les compléments doivent, dans la mesure du possible, utiliser les API REST plutôt que les services Web Exchange. 

**Jetons REST**

Quand un jeton REST est demandé (`options.isRest = true`), le jeton fourni ne permet pas d’authentifier les appels des services Web Exchange. Le jeton peut seulement accéder à l’élément actif et à ses pièces jointes en lecture seule, sauf si l’autorisation [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) est spécifiée dans le manifeste du complément. Si l’autorisation `ReadWriteMailbox` est spécifiée, le jeton fourni accorde un accès en lecture/écriture au courrier, au calendrier et aux contacts, ainsi que la possibilité d’envoyer des messages.

Le complément doit utiliser la propriété `restUrl` pour déterminer l’URL à utiliser pendant les appels de l’API REST.

**Jetons EWS**

Lorsque un jeton EWS est demandé (`options.isRest = false`), le jeton fourni ne permet pas d’authentifier les appels de l’API REST. Le jeton à une étendue limitée à l’accès à l’élément actif.

Le complément doit utiliser la propriété `ewsUrl` pour déterminer l’URL à utiliser pendant les appels EWS.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Attributs| Description|
|---|---|---|---|
| `options` | Objet | &lt;facultatif&gt; | Littéral d’objet contenant une ou plusieurs des propriétés suivantes. |
| `options.isRest` | Boolean |  &lt;facultatif&gt; | Détermine si le jeton fourni est utilisé pour les API REST Outlook ou les services Web Exchange. La valeur par défaut est `false`. |
| `options.asyncContext` | Objet |  &lt;facultatif&gt; | Toute donnée d'état qui est passée à la méthode asynchrone. |
|`callback`| fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.5 |
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition et lecture|

##### <a name="example"></a>Exemple

```js
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Obtient une chaîne qui contient un jeton servant à obtenir une pièce jointe ou un élément à partir d’un Exchange Server.

La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir du Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.

Vous pouvez transmettre le jeton et un identificateur de pièce jointe ou un identificateur d’élément à un système tiers. Le système tiers utilise le jeton comme jeton d’autorisation au porteur pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des services Web Exchange, pour retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Votre application doit disposer de l’autorisation **ReadItem** spécifiée dans son manifeste, pour pouvoir appeler la méthode `getCallbackTokenAsync` en mode lecture.

En mode composition, vous devez appeler la méthode [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) permettant d’obtenir un identificateur de l’élément à transmettre à la méthode `getCallbackTokenAsync`. Votre application doit disposer des autorisations **ReadWriteItem** pour appeler la méthode `saveAsync`.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Attributs| Description|
|---|---|---|---|
|`callback`| fonction||Une fois la méthode exécutée, la fonction transmise au paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult). Le jeton est fourni sous forme de chaîne dans la propriété `asyncResult.value`.|
|`userContext`| Objet| &lt;facultatif&gt;|Toute donnée d'état qui est passée à la méthode asynchrone.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.3|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition et lecture|

##### <a name="example"></a>Exemple

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync (rappel, [userContext])

Obtient un jeton identifiant l’utilisateur et le complément Office.

La méthode `getUserIdentityTokenAsync` retourne un jeton que vous pouvez utiliser pour identifier et [authentifier le complément et l’utilisateur avec un système de tierce partie](https://docs.microsoft.com/outlook/add-ins/authentication).

##### <a name="parameters"></a>Paramètres :

|Name| Type| Attributs| Description|
|---|---|---|---|
|`callback`| fonction||Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le jeton est fourni sous la forme d'une chaîne dans la propriété `asyncResult.value`.|
|`userContext`| Objet| &lt;facultatif&gt;|Toute donnée d'état qui est passée à la méthode asynchrone.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync (données, rappel [userContext])

Effectue une demande asynchrone à un service Exchange Web Services (EWS) sur l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur.

> [!NOTE]
> Cette méthode n’est pas pris en charge dans les scénarios suivants.
> - Dans Outlook pour iOS ou Outlook pour Android
> - Lorsque le complément est chargé dans une boîte aux lettres Gmail
> 
> Dans ces cas, les compléments doivent [utiliser l’API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) à la place pour accéder à la boîte aux lettres de l’utilisateur.

La méthode `makeEwsRequestAsync` envoie une demande EWS à Exchange de la part du complément. Voir [Appeler des services web depuis un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) pour une liste des opérations EWS prises en charge, .

Vous ne pouvez pas demander Folder Associated Items avec la méthode `makeEwsRequestAsync`.

La demande XML doit spécifier l’encodage UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et sur les opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, voir la page relative aux [Indiquer des autorisations pour l'accès du complément de messagerie à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> L’administrateur du serveur doit définir `OAuthAuthentication` à true dans le dossier EWS du serveur d’accès client, pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.

##### <a name="version-differences"></a>Différences entre versions

Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie s'exécutant dans des versions d’Outlook inférieures à la version 15.0.4535.1004, vous devez définir la valeur d’encodage à `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Vous n’avez pas besoin de définir la valeur d’encodage quand votre application de messagerie s’exécute dans Outlook sur le web. Vous pouvez déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le web en utilisant la propriété mailbox.diagnostics.hostName. Vous pouvez déterminer quelle version d’Outlook est exécutée en utilisant la propriété mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Paramètres :

|Name| Type| Attributs| Description|
|---|---|---|---|
|`data`| String||La demande EWS.|
|`callback`| fonction||Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le résultat XML de l’appel EWS est fourni comme une chaîne dans la propriété `asyncResult.value`. Si le résultat dépasse 1 Mo en taille, un message d’erreur est retourné à la place.|
|`userContext`| Objet| &lt;facultatif&gt;|Toute donnée d'état qui est passée à la méthode asynchrone.|

##### <a name="requirements"></a>Conditions requises

|Condition| Valeur|
|---|---|
|[Version minimale de l’ensemble de conditions requises de boîte aux lettres](/javascript/office/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau minimal d’autorisation](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

L’exemple suivant appelle `makeEwsRequestAsync` pour utiliser l’opération `GetItem` pour obtenir l’objet d’un élément.

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