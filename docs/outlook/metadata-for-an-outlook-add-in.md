---
title: Obtenir et définir des métadonnées dans un complément Outlook
description: Vous pouvez gérer les données personnalisées dans votre complément Outlook en utilisant les paramètres d’itinérance ou propriétés personnalisées.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: ceed27cc5c0d479ac67a0497e78e971498365e6f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58939280"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a>Obtenir et définir des métadonnées de complément pour un complément Outlook

Vous pouvez gérer les données personnalisées dans votre complément Outlook en utilisant une des solutions suivantes :

- Les paramètres d’itinérance, qui permettent de gérer des données personnalisées pour la boîte aux lettres d’un utilisateur.
- Les propriétés personnalisées, qui permettent de gérer des données personnalisées pour un élément de boîte aux lettres d’un utilisateur.

Ces deux méthodes donnent accès aux données personnalisées auxquelles seul votre complément Outlook a accès, mais chaque méthode stocke les données de façon distincte. Autrement dit, les propriétés personnalisées n’ont pas accès aux données stockées par le biais des paramètres d’itinérance et inversement. Les données sont stockées sur le serveur de la boîte aux lettres et sont accessibles dans les sessions Outlook ultérieures sur tous les formats pris en charge par le complément.

## <a name="custom-data-per-mailbox-roaming-settings"></a>Données personnalisées par boîte aux lettres : paramètres d’itinérance

Vous pouvez indiquer des données propres à la boîte aux lettres Exchange d’un utilisateur, à l’aide de l’objet [RoamingSettings](/javascript/api/outlook/office.RoamingSettings), telles que les préférences et les données personnelles de l’utilisateur. Votre complément de messagerie peut accéder aux paramètres d’itinérance lorsqu’il est en itinérance sur un appareil pour lequel il a été conçu (ordinateur, tablette ou smartphone).

Les modifications apportées à ces données sont stockées dans une copie en mémoire de ces paramètres pour la session Outlook en cours. Vous devez explicitement enregistrer tous les paramètres d’itinérance après les avoir mis à jour afin qu’ils soient disponibles lors de la prochaine ouverture de votre complément, sur le même appareil ou sur un autre appareil pris en charge.


### <a name="roaming-settings-format"></a>Format des paramètres d’itinérance

Les données dans un objet **RoamingSettings** sont stockées sous forme d’une chaîne sérialisée JavaScript Object Notation (JSON). 

Voici un exemple de structure, en supposant qu’il y a trois paramètres d’itinérance définis nommés `add-in_setting_name_0`, `add-in_setting_name_1` et `add-in_setting_name_2`.


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a>Chargement des paramètres d’itinérance

Un complément de messagerie charge généralement les paramètres d’itinérance dans le gestionnaire d’événements [Office.initialize](/javascript/api/office#Office_initialize_reason_). L’exemple de code JavaScript suivant montre comment charger les paramètres d’itinérance existants et obtenir les valeurs de 2 paramètres, **customerName** et **customerBalance**.


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a>Création ou affectation d’un paramètre d’itinérance

Pour faire suite à l’exemple précédent, la fonction JavaScript suivante, `setAddInSetting`, montre comment utiliser la méthode [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) pour définir un paramètre nommé `cookie` avec la date du jour, et conserver les données en utilisant la méthode [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveAsync_callback_) pour réenregistrer tous les paramètres d’itinérance sur le serveur.

La méthode crée le paramètre si le paramètre n’existe pas déjà et affecte le paramètre `set` à la valeur spécifiée. La `saveAsync` méthode enregistre les paramètres d’itinérance de manière asynchrone. Cet exemple de code transmet une méthode de rappel, à « When the asynchronous call finishes » (Lorsque l’appel asynchrone se termine), est appelée à l’aide d’un `saveMyAddInSettingsCallback` `saveAsync`  `saveMyAddInSettingsCallback` paramètre, _asyncResult_. Ce paramètre est un objet [AsyncResult](/javascript/api/office/office.asyncresult) qui contient le résultat des détails relatifs à l’appel asynchrone. Vous pouvez utiliser le paramètre facultatif _userContext_ pour transmettre des informations d’état de l’appel asynchrone à la fonction de rappel.

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a>Suppression d’un paramètre d’itinérance

Toujours dans le prolongement des exemples précédents, la fonction JavaScript suivante,  `removeAddInSetting`, illustre l’utilisation de la méthode [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove_name_) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a>Données personnalisées par élément dans une boîte aux lettres : propriétés personnalisées

Vous pouvez spécifier les données propres à un élément dans la boîte aux lettres de l’utilisateur à l’aide de l’objet [CustomProperties](/javascript/api/outlook/office.CustomProperties). Par exemple, votre complément de messagerie peut catégoriser certains messages et noter la catégorie à l’aide d’une propriété personnalisée`messageCategory`. Si votre complément de messagerie crée des rendez-vous à partir de suggestions de réunion dans un message, vous pouvez utiliser une propriété personnalisée pour suivre chacun de ces rendez-vous. Cela garantit que si l’utilisateur ouvre à nouveau le message, votre complément de messagerie ne propose pas de créer le rendez-vous une seconde fois.

Comme pour les paramètres d’itinérance, les modifications apportées aux propriétés personnalisées sont stockées dans des copies en mémoire des propriétés de la session Outlook en cours. Pour vous assurer que les propriétés personnalisées seront disponibles dans la prochaine session, utilisez [CustomProperties.saveAsync](/javascript/api/outlook/office.customproperties#saveAsync_callback__asyncContext_).

Ces propriétés personnalisées spécifiques à un élément et spécifiques au add-in sont accessibles uniquement à l’aide de `CustomProperties` l’objet. Ces propriétés sont différentes des propriétés [UserProperties](/office/vba/api/Outlook.UserProperties) personnalisées basées sur MAPI dans le modèle objet Outlook et des propriétés étendues dans Exchange Web Services (EWS). Vous ne pouvez pas accéder directement `CustomProperties` à l’aide Outlook modèle objet, EWS ou REST. Pour savoir comment accéder à l’aide d’EWS ou rest, voir la section Obtenir des propriétés personnalisées à l’aide `CustomProperties` [d’EWS ou REST](#get-custom-properties-using-ews-or-rest).

### <a name="using-custom-properties"></a>Utilisation de propriétés personnalisées

Avant de pouvoir utiliser les propriétés personnalisées, vous devez les charger en appelant la méthode [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods). Après avoir créé le conteneur de propriétés, vous pouvez utiliser les méthodes [Définir](/javascript/api/outlook/office.customproperties#set_name__value_) et [Obtenir](/javascript/api/outlook/office.customproperties) pour ajouter et récupérer des propriétés personnalisées. Vous devez utiliser la méthode[saveAsync](/javascript/api/outlook/office.customproperties#saveAsync_callback__asyncContext_) pour enregistrer les modifications que vous apportez au conteneur de propriétés.


 > [!NOTE]
 > Comme Outlook sur Mac ne met pas en cache les propriétés personnalisées, si le réseau de l’utilisateur tombe en panne, les compléments de messagerie dans Outlook sur Mac ne pourront pas accéder à leurs propriétés personnalisées.


### <a name="custom-properties-example"></a>Exemple de propriétés personnalisées


L’exemple suivant illustre un ensemble simplifié des méthodes pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément qui utilise des propriétés personnalisées.

Cet exemple inclut les méthodes suivantes.


- [Office.initialize](/javascript/api/office#Office_initialize_reason_) -- Initialise le complément et charge le conteneur de propriétés personnalisées depuis le serveur Exchange.

- **customPropsCallback** -- Obtient le conteneur de propriétés personnalisées qui est renvoyé depuis le serveur et l’enregistre pour une utilisation ultérieure.

- **updateProperty** -- Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur.

- **removeProperty** -- Supprime une propriété spécifique à partir du conteneur de propriétés, puis enregistre la suppression sur le serveur.


```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="get-custom-properties-using-ews-or-rest"></a>Obtenir des propriétés personnalisées à l’aide de EWS ou REST

Pour obtenir **CustomProperties** à l’aide de EWS ou REST, vous devez commencer par déterminer le nom de sa base propriété étendue MAPI. Vous pouvez ensuite obtenir cette propriété de la même façon que vous pouviez obtenir toute propriété étendue de base MAPI.

#### <a name="how-custom-properties-are-stored-on-an-item"></a>Comment les propriétés personnalisées sont stockées sur un élément

Les propriétés personnalisées définies par un complément ne sont pas équivalentes aux propriétés de base MAPI normales. Les API de votre add-in sérialisent tous vos modules en tant que charge utile JSON, puis les enregistrent dans une seule propriété étendue basée sur MAPI dont le nom est ( est l’ID de votre `CustomProperties` `cecp-<app-guid>` `<app-guid>` add-in) et le GUID du jeu de propriétés est `{00020329-0000-0000-C000-000000000046}` . (Pour plus d’informations sur cet objet, voir[MS-OXCEXT 2.2.5 Propriétés d’Application de messagerie Personnalisées](/openspecs/exchange_server_protocols/ms-oxcext/4cf1da5e-c68e-433e-a97e-c45625483481).) Vous pouvez ensuite utiliser EWS ou REST pour obtenir cette propriété basée MAPI.

#### <a name="get-custom-properties-using-ews"></a>Obtenir des propriétés personnalisées à l’aide de EWS

Votre add-in de messagerie peut obtenir la propriété étendue basée sur MAPI à l’aide de l’opération `CustomProperties` [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) EWS. Accès côté serveur à l’aide d’un jeton de rappel ou côté client à l’aide de la méthode `GetItem` [mailbox.makeEwsRequestAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) Dans la demande, spécifiez la propriété basée sur MAPI dans son jeu de propriétés à l’aide des détails fournis dans la section précédente Comment les `GetItem` `CustomProperties` propriétés personnalisées sont stockées [sur un élément](#how-custom-properties-are-stored-on-an-item).

L’exemple suivant montre comment obtenir un élément et ses propriétés personnalisées.

> [!IMPORTANT]
> Dans l’exemple suivant, remplacez `<app-guid>` par l’ID de votre complément.

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

Vous pouvez également obtenir plus de propriétés personnalisées si vous les spécifiez dans la chaîne de demande, comme les autres éléments [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri).

#### <a name="get-custom-properties-using-rest"></a>Obtenir des propriétés personnalisées à l’aide de REST

Dans votre complément, vous pouvez construire votre requête REST contre les messages et événements pour obtenir ceux qui déjà ont des propriétés personnalisées. Dans la requête, spécifiez la propriété basée MAPI **CustomProperties** dans son ensemble de propriété à l’aide des informations fournies dans la section précédente [Comment les propriétés personnalisées sont stockées sur un élément](#how-custom-properties-are-stored-on-an-item).

L’exemple suivant montre comment obtenir tous les événements ayant des propriétés personnalisées définies par votre complément et vous assurer que la réponse inclut la valeur de la propriété pour vous permettre d’appliquer une logique de filtrage.

> [!IMPORTANT]
> Dans l’exemple suivant, remplacez `<app-guid>` avec l’ID de votre complément.

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

Pour plus exemples qui utilisent REST pour obtenir les propriétés étendues à valeur unique base MAPI, voir [Obtenir singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0&preserve-view=true).

L’exemple suivant montre comment obtenir un élément et ses propriétés personnalisées. Dans la fonction de rappel pour la méthode `done`, `item.SingleValueExtendedProperties` contient la liste des propriétés personnalisées demandées.

> [!IMPORTANT]
> Dans l’exemple suivant, remplacez `<app-guid>` par l’ID de votre complément.

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## <a name="see-also"></a>Voir aussi

- [Vue d’ensemble de la propriété MAPI](/office/client-developer/outlook/mapi/mapi-property-overview)
- [Présentation des propriétés Outlook](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [Utilisation des API REST Outlook d’un complément Outlook](use-rest-api.md)
- [Appeler des services web à partir d’un complément Outlook](web-services.md)
- [Les propriétés et les propriétés étendues dans EWS dans Exchange](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [Jeux de propriétés et de réponse des formes dans EWS dans Exchange](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)