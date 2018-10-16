
# <a name="mailbox"></a>mailbox

### [Office](Office.md)[.context](Office.context.md). mailbox

Donne accès au modèle objet du complément Outlook pour Microsoft Outlook et Microsoft Outlook sur le web.

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale de l’ensemble des conditions requises de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| Restreint|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

### <a name="namespaces"></a>Espaces de noms

[diagnostics](Office.context.mailbox.diagnostics.md) : fournit des informations de diagnostic à un complément Outlook.

[item](Office.context.mailbox.item.md) : fournit des méthodes et des propriétés pour accéder à un message ou un rendez-vous dans un complément Outlook.

[userProfile](Office.context.mailbox.userProfile.md) : fournit des informations sur l’utilisateur dans un complément Outlook.

### <a name="members"></a>Membres

#### <a name="ewsurl-string"></a>ewsUrl :String

Obtient l’URL du point de terminaison des services Web Exchange  (EWS) pour ce compte de messagerie. Mode lecture uniquement.

> [!NOTE]
> Ce membre n’est pas pris en charge par Outlook pour iOS ou Outlook pour Android.

La valeur `ewsUrl` peut être utilisée par un service distant pour effectuer des appels EWS vers la boîte aux lettres de l’utilisateur. Par exemple, vous pouvez créer un service distant pour [obtenir les pièces jointes de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

##### <a name="type"></a>Type :

*   Chaîne

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lecture|

### <a name="methods"></a>Méthodes

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime"></a>convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}

Obtient un dictionnaire contenant des informations d’heure dans l’heure locale du client.

Les dates et heures utilisées par une application de messagerie pour Outlook ou Outlook Web App peuvent utiliser des fuseaux horaires différents. Outlook utilise le fuseau horaire de l’ordinateur client, Outlook Web App utilise le fuseau horaire défini dans le Centre d’administration Exchange (CAE). Vous devez gérer les valeurs de date et d’heure de telle sorte que les valeurs que vous affichez sur l’interface utilisateur soient toujours cohérentes avec le fuseau horaire auquel l’utilisateur s'attend.

Si l’application de messagerie s'exécute dans Outlook, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire de l’ordinateur client. Si l’application de messagerie s’exécute dans Outlook Web App, la méthode `convertToLocalClientTime` retournera un objet dictionnaire dont les valeurs seront définies pour le fuseau horaire spécifié dans le CAE.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`timeValue`| Date|Un objet Date|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="returns"></a>Retourne :

Type : [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)

####  <a name="converttoutcclienttimeinput--date"></a>convertToUtcClientTime(input) → {Date}

Obtient un objet Date à partir d’un dictionnaire contenant des informations d’heure.

La méthode `convertToUtcClientTime` convertit un dictionnaire contenant une date et une heure locales en objet Date avec les valeurs correctes pour la date et l’heure locales.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`input`| [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)|Valeur en heure locale à convertir.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
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

La méthode `displayAppointmentForm` ouvre un rendez-vous de calendrier existant dans une nouvelle fenêtre sur le bureau ou dans une boîte de dialogue sur les équipements mobiles.

Dans Outlook pour Mac, vous pouvez utiliser cette méthode pour afficher un seul rendez-vous qui ne fait pas partie d’une série périodique, ou le rendez-vous principal d’une série périodique, mais vous ne pouvez pas afficher une instance de la série. Cela est dû au fait que, dans Outlook pour Mac, vous ne pouvez pas accéder aux propriétés (y compris l’ID d’élément) des instances d’une série périodique.

Dans Outlook Web App, cette méthode ouvre le formulaire spécifié seulement si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.

Si l’identificateur d’élément indiqué n’identifie pas un rendez-vous existant, un volet vierge s’ouvre sur l’ordinateur ou l’appareil client, et aucun message d’erreur n'est retourné.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`itemId`| Chaîne|L'identificateur EWS (services web Exchange) pour un rendez-vous de calendrier existant.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
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

Dans Outlook Web App, cette méthode n’ouvre le formulaire indiqué que si le corps du formulaire comprend un nombre de caractères inférieur ou égal à 32 Ko.

Si l’identificateur d’élément indiqué n’identifie pas un message existant, aucun message ne sera affiché sur l’ordinateur client, et aucun message d’erreur ne sera retourné.

N’utilisez pas la méthode `displayMessageForm` avec un `itemId` qui représente un rendez-vous. Utilisez la méthode `displayAppointmentForm` pour afficher un rendez-vous existant, et `displayNewAppointmentForm` pour afficher un formulaire pour créer un nouveau rendez-vous.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
|`itemId`| Chaîne|Identificateur EWS (services web Exchange) pour un message existant.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
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

Dans Outlook Web App et OWA pour les appareils, cette méthode affiche toujours un formulaire avec un champ participants. Si vous n'indiquez aucun participant dans les arguments d’entrée, la méthode affiche un formulaire avec un bouton **Enregistrer**. Si vous avez indiqué des participants, le formulaire inclura les participants et un bouton **Envoyer**.

Dans le client riche Outlook et Outlook RT, si vous indiquez des participants ou des ressources dans les paramètres `requiredAttendees`, `optionalAttendees`, ou `resources`, cette méthode affiche un formulaire de réunion avec un bouton **Envoyer**. Si vous ne n'indiquez aucun destinataire, cette méthode affiche un formulaire de rendez-vous avec un bouton **Enregistrer et fermer**.

Si l’un des paramètres dépasse les limites de taille indiquées, ou si un nom de paramètre inconnu est indiqué, une exception est levée.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Description|
|---|---|---|
| `parameters` | Objet | Un dictionnaire de paramètres décrivant le nouveau rendez-vous. |
| `parameters.requiredAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt; | Un tableau de chaînes contenant les adresses de messagerie ou un tableau contenant un objet `EmailAddressDetails` pour chacun des participants requis pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées. |
| `parameters.optionalAttendees` | Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt; | Un tableau de chaînes contenant les adresses de messagerie ou tableau contenant un objet `EmailAddressDetails` pour chacun des participants facultatifs du rendez-vous. Le tableau est limité à un maximum de 100 entrées. |
| `parameters.start` | Date | Un objet `Date` indiquant la date et l’heure du début du rendez-vous. |
| `parameters.end` | Date | Un objet `Date` indiquant la date et l’heure de la fin du rendez-vous. |
| `parameters.location` | Chaîne | Une chaîne contenant le lieu du rendez-vous. La chaîne est limitée à un maximum de 255 caractères. |
| `parameters.resources` | Array.&lt;String&gt; | Un tableau de chaînes contenant les ressources requises pour le rendez-vous. Le tableau est limité à un maximum de 100 entrées. |
| `parameters.subject` | Chaîne | Une chaîne contenant l’objet du rendez-vous. La chaîne est limitée à un maximum de 255 caractères. |
| `parameters.body` | Chaîne | Le corps du rendez-vous. Le contenu du corps est limité à une taille maximale de 32 Ko. |

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale de l’ensemble des conditions requises de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
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

#### <a name="getcallbacktokenasynccallback-usercontext"></a>getCallbackTokenAsync(callback, [userContext])

Obtient une chaîne qui contient un jeton utilisé pour obtenir une pièce jointe ou un élément à partir d’un Exchange Server.

La méthode `getCallbackTokenAsync` effectue un appel asynchrone pour obtenir un jeton opaque à partir du Exchange Server qui héberge la boîte aux lettres de l’utilisateur. La durée de vie du jeton de rappel est de 5 minutes.

Vous pouvez passer le jeton et un identificateur de pièce jointe ou un identificateur d’élément à un système de tierce partie. Celui-ci utilise le jeton en tant que jeton d’autorisation du porteur pour appeler l’opération [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) ou [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) des Exchange Web Services (EWS) pour retourner une pièce jointe ou un élément. Par exemple, vous pouvez créer un service distant pour [obtenir des pièces jointes à partir de l’élément sélectionné](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).

Votre application doit disposer de l’autorisation **ReadItem** indiquée dans son manifeste pour appeler la méthode `getCallbackTokenAsync`.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| fonction||Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le jeton est fourni sous la forme d’une chaîne dans la propriété `asyncResult.value`.|
|`userContext`| Objet| &lt;facultatif&gt;|Toute donnée d’état qui est passée à la méthode asynchrone.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Lecture|

##### <a name="example"></a>Exemple

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a>getUserIdentityTokenAsync(callback, [userContext])

Obtient un jeton identifiant l’utilisateur et le complément Office.

La méthode `getUserIdentityTokenAsync` retourne un jeton que vous pouvez utiliser pour identifier et [authentifier le complément et l’utilisateur avec un système de tiers](https://docs.microsoft.com/outlook/add-ins/authentication).

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`callback`| fonction||Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un seul paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le jeton est fourni sous la forme d’une chaîne dans la propriété `asyncResult.value`.|
|`userContext`| Objet| &lt;facultatif&gt;|Toute donnée d’état qui est passée à la méthode asynchrone.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale de l’ensemble des conditions requises de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadItem|
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

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a>makeEwsRequestAsync(data, callback, [userContext])

Effectue une demande asynchrone à un service Exchange Web Services (EWS) sur l'Exchange Server qui héberge la boîte aux lettres de l’utilisateur.

> [!NOTE]
> Cette méthode n’est pas prise en charge dans les scénarios suivants.
> - Dans Outlook pour iOS ou Outlook pour Android
> - Lorsque le complément est chargé dans une boîte aux lettres Gmail
> 
> Dans ces cas, les compléments doivent [utiliser l’API REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) à la place pour accéder à la boîte aux lettres de l’utilisateur.

La `makeEwsRequestAsync` méthode envoie une demande EWS à Exchange, de la part du complément. Pour obtenir la liste des opérations EWS prises en charge, voir [appeler des services web à partir d’un complément Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) .

Avec la méthode `makeEwsRequestAsync` , vous ne pouvez pas demander des éléments associés à un dossier.

La demande XML doit spécifier l’encodage UTF-8.

```
<?xml version="1.0" encoding="utf-8"?>
```

Votre complément doit disposer de l’autorisation **ReadWriteMailbox** pour utiliser la méthode `makeEwsRequestAsync`. Pour plus d’informations sur l’utilisation de l’autorisation **ReadWriteMailbox** et sur les opérations EWS que vous pouvez appeler avec la méthode `makeEwsRequestAsync`, voir la page [Indiquer des autorisations pour l'accès du complément de messagerie à la boîte aux lettres de l’utilisateur](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).

> [!NOTE]
> L’administrateur du serveur doit définir `OAuthAuthentication` à true dans le dossier EWS du serveur d’accès client, pour permettre à la méthode `makeEwsRequestAsync` d’effectuer des demandes EWS.

##### <a name="version-differences"></a>Différences entre versions

Lorsque vous utilisez la méthode `makeEwsRequestAsync` dans les applications de messagerie s'exécutant dans des versions d’Outlook antérieures à la version 15.0.4535.1004, vous devez définir la valeur d’encodage sur `ISO-8859-1`.

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

Vous n’avez pas besoin de définir la valeur d’encodage quand votre application de messagerie s’exécute dans Outlook sur le Web. Vous pouvez déterminer si votre application de messagerie s’exécute dans Outlook ou Outlook sur le Web en utilisant la propriété mailbox.diagnostics.hostName. Vous pouvez déterminer quelle version d’Outlook est exécutée en utilisant la propriété mailbox.diagnostics.hostVersion.

##### <a name="parameters"></a>Paramètres :

|Nom| Type| Attributs| Description|
|---|---|---|---|
|`data`| Chaîne||La demande EWS.|
|`callback`| fonction||Quand la méthode se termine, la fonction passée dans le paramètre `callback` est appelée avec un simple paramètre, `asyncResult`, qui est un objet [`AsyncResult`](/javascript/api/office/office.asyncresult).<br/><br/>Le résultat XML de l’appel EWS est fourni sous forme d'une chaîne dans la propriété `asyncResult.value` . Si le résultat dépasse la taille d'1 Mo, un message d’erreur est renvoyé à la place.|
|`userContext`| Objet| &lt;facultatif&gt;|Toute donnée d’état qui est passée à la méthode asynchrone.|

##### <a name="requirements"></a>Conditions requises

|Condition requise| Valeur|
|---|---|
|[Version minimale des exigences de boîte aux lettres](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| 1.0|
|[Niveau d’autorisation minimal](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| ReadWriteMailbox|
|[Mode Outlook applicable](https://docs.microsoft.com/outlook/add-ins/#extension-points)| Composition ou lecture|

##### <a name="example"></a>Exemple

L’exemple suivant appelle `makeEwsRequestAsync` pour utiliser l’opération `GetItem` et obtenir l’objet d’un élément.

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