---
title: Gérer l’État et les paramètres d’un complément Outlook
description: Découvrez comment faire persister l’État et les paramètres des compléments pour un complément Outlook.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 7d981107da68c329d209834059bfac494d6ccae4
ms.sourcegitcommit: 4079903c3cc45b7d8c041509a44e9fc38da399b1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 03/11/2020
ms.locfileid: "42596647"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>Gérer l’État et les paramètres d’un complément Outlook

> [!NOTE]
> Consultez la section relative à la conservation de l' [État et des paramètres des compléments](../develop/persisting-add-in-state-and-settings.md) dans la section **concepts principaux** de cette documentation avant de lire cet article.

Pour les compléments Outlook, l’API JavaScript pour Office fournit des objets [RoamingSettings](/javascript/api/outlook/office.roamingsettings) et [CustomProperties](/javascript/api/outlook/office.customproperties) pour enregistrer l’état des compléments entre les sessions, comme décrit dans le tableau suivant. Dans tous les cas, les valeurs de paramètre enregistrées sont associées à l’[ID](../reference/manifest/id.md) du complément qui les a créées.

|**Objet**|**Emplacement de stockage**|
|:-----|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Boîte aux lettres de serveur Exchange de l’utilisateur où le complément est installé. Comme ces paramètres sont stockés dans la boîte aux lettres de serveur de l’utilisateur, ils sont itinérants et accessibles par le complément lorsqu’il s’exécute dans le contexte d’une application hôte cliente ou d’un navigateur pris en charge accédant à la boîte aux lettres de cet utilisateur.<br/><br/> Seul le complément qui a créé les paramètres d’itinérance du complément Outlook peut y accéder, et uniquement dans la boîte aux lettres où le complément est installé.|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Élément de message, de rendez-vous ou de demande de réunion qu’utilise le complément. Seul le complément qui a créé les propriétés personnalisées d’élément de complément Outlook peut y accéder, et uniquement dans l’élément où elles sont enregistrées.|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Enregistrement des paramètres en tant que paramètres d’itinérance dans la boîte aux lettres de l’utilisateur pour les compléments Outlook

Un complément Outlook peut utiliser l’objet [RoamingSettings](/javascript/api/outlook/office.roamingsettings) pour enregistrer les données de paramètres et d’état du complément propres à la boîte aux lettres de l’utilisateur. Seul ce complément Outlook peut accéder aux données pour le compte de l’utilisateur qui exécute le complément. Les données sont stockées dans la boîte aux lettres Exchange Server de l’utilisateur et sont accessibles lorsque cet utilisateur se connecte à son compte et exécute le complément Outlook.

### <a name="loading-roaming-settings"></a>Chargement des paramètres d’itinérance

Un complément Outlook charge généralement les paramètres d’itinérance dans le gestionnaire d’événements [Office.initialize](/javascript/api/office). L’exemple de code JavaScript suivant explique comment charger des paramètres d’itinérance existants.

```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a>Création ou affectation d’un paramètre d’itinérance

Pour faire suite à l’exemple précédent, la fonction `setAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) pour définir ou mettre à jour un paramètre nommé `cookie` avec la date du jour. Elle réenregistre ensuite tous les paramètres d’itinérance sur le serveur Exchange avec la méthode [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).

```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

La méthode **saveAsync** enregistre les paramètres d’itinérance de manière asynchrone et admet une fonction de rappel facultative. Cet exemple de code transmet une fonction de rappel nommée `saveMyAppSettingsCallback` à la méthode **saveAsync**. Lors du renvoi de l’appel asynchrone, le paramètre _asyncResult_ de la fonction `saveMyAppSettingsCallback` fournit un accès à un objet [AsyncResult](/javascript/api/outlook) que vous pouvez utiliser pour déterminer le succès ou l’échec de l’opération avec la propriété**AsyncResult.status**.

### <a name="removing-a-roaming-setting"></a>Suppression d’un paramètre d’itinérance

Toujours dans le prolongement des exemples précédents, la fonction  `removeAppSetting` suivante montre comment utiliser la méthode [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) pour supprimer le paramètre `cookie` et réenregistrer tous les paramètres d’itinérance sur le serveur Exchange.

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Enregistrement des paramètres par élément pour les compléments Outlook en tant que propriétés personnalisées

Les propriétés personnalisées permettent à votre complément Outlook de stocker des informations sur un élément qu’il utilise. Par exemple, si votre complément Outlook crée un rendez-vous à partir d’une suggestion de réunion dans un message, vous pouvez utiliser des propriétés personnalisées pour stocker le fait que la réunion a été créée. Cela garantit que si le message est rouvert, votre complément Outlook ne propose pas de recréer le rendez-vous.

Pour pouvoir utiliser des propriétés personnalisées pour un élément de message, de rendez-vous ou de demande de réunion particulier, vous devez charger les propriétés en mémoire en appelant la méthode [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) de l’objet **Item**. Si des propriétés personnalisées sont déjà définies pour l’élément actuel, elles sont chargées à ce moment à partir du serveur Exchange. Après avoir chargé les propriétés, vous pouvez utiliser les méthodes [set](/javascript/api/outlook/office.customproperties#set-name--value-) et [get](/javascript/api/outlook/office.roamingsettings) de l’objet **CustomProperties** pour ajouter, mettre à jour et récupérer des propriétés en mémoire. Pour enregistrer les modifications que vous avez apportées aux propriétés personnalisées de l’élément, vous devez utiliser la méthode [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) pour conserver les modifications de l’élément sur le serveur Exchange.

### <a name="custom-properties-example"></a>Exemple de propriétés personnalisées

L’exemple suivant illustre un ensemble simplifié des fonctions pour un complément Outlook qui utilise des propriétés personnalisées. Vous pouvez utiliser cet exemple comme point de départ pour votre complément Outlook qui utilise des propriétés personnalisées. 

Un complément Outlook qui utilise ces fonctions récupère toutes les propriétés personnalisées en appelant la méthode **get** sur la variable `_customProps`, comme le montre l’exemple suivant.

```js
var property = _customProps.get("propertyName");
```

Cet exemple inclut les fonctions suivantes :

|**Nom de la fonction**|**Description**|
|:-----|:-----|
| `Office.initialize`|Initialise le complément et charge les propriétés personnalisées pour l’élément actuel à partir du serveur Exchange.|
| `customPropsCallback`|Obtient les propriétés personnalisées retournées du serveur Exchange et les enregistre pour une utilisation ultérieure.|
| `updateProperty`|Définit ou met à jour une propriété spécifique, puis enregistre la modification sur le serveur Exchange.|
| `removeProperty`|Supprime une propriété spécifique, puis fait persister la suppression sur le serveur Exchange.|
| `saveCallback`|Rappel pour les appels à la méthode**saveAsync** dans les fonctions`updateProperty` et `removeProperty`.|

```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method.
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

## <a name="see-also"></a>Voir aussi

- [Conservation de l’état et des paramètres des compléments](../develop/persisting-add-in-state-and-settings.md)
- [Initialiser votre complément Office](../develop/initialize-add-in.md)