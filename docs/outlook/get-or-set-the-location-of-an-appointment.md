---
title: Obtenir ou définir le lieu de rendez-vous dans un complément
description: Découvrez comment obtenir ou définir l’heure d’un rendez-vous à partir d’un complément Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 5669f656348465baabb3e684b359261024a509ca
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671834"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook

L Office API JavaScript fournit des propriétés et des méthodes pour gérer l’emplacement d’un rendez-vous que l’utilisateur compose. Actuellement, deux propriétés fournissent l’emplacement d’un rendez-vous :

- [item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API de base qui vous permet d’obtenir et de définir l’emplacement.
- [item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API améliorée qui vous permet d’obtenir et de définir l’emplacement, et inclut la spécification du [type d’emplacement.](/javascript/api/outlook/office.mailboxenums.locationtype) Le type est `LocationType.Custom` si vous définissez l’emplacement à l’aide `item.location` de .

Le tableau suivant répertorie les API d’emplacement et les modes (c’est-à-dire, composer ou lire) où elles sont disponibles.

| API | Modes de rendez-vous applicables |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#location) | Attendee/Read |
| [item.location.getAsync](/javascript/api/outlook/office.location#getAsync_options__callback_) | Organisateur/Composition |
| [item.location.setAsync](/javascript/api/outlook/office.location#setAsync_location__options__callback_) | Organisateur/Composition |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#getAsync_options__callback_) | Organisateur/Composition,<br>Attendee/Read |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#addAsync_locationIdentifiers__options__callback_) | Organisateur/Composition |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#removeAsync_locationIdentifiers__options__callback_) | Organisateur/Composition |

Pour utiliser les méthodes disponibles uniquement pour composer des add-ins, configurez le manifeste du add-in pour activer le module en mode Organisateur/Composition. Pour [plus d Outlook,](compose-scenario.md) voir Créer des Outlook pour les formulaires de composition.

## <a name="use-the-enhancedlocation-api"></a>Utiliser `enhancedLocation` l’API

Vous pouvez utiliser `enhancedLocation` l’API pour obtenir et définir l’emplacement d’un rendez-vous. Le champ Emplacement prend en charge plusieurs emplacements et, pour chaque emplacement, vous pouvez définir le nom complet, le type et l’adresse e-mail de la salle de conférence (le cas échéant). Voir [LocationType pour](/javascript/api/outlook/office.mailboxenums.locationtype) les types d’emplacement pris en charge.

### <a name="add-location"></a>Ajouter un emplacement

L’exemple suivant montre comment ajouter un emplacement en appelant [addAsync](/javascript/api/outlook/office.enhancedlocation#addAsync_locationIdentifiers__options__callback_) sur [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedLocation).

```js
var item;
var locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>Obtenir un emplacement

L’exemple suivant montre comment obtenir l’emplacement en appelant [getAsync](/javascript/api/outlook/office.enhancedlocation#getAsync_options__callback_) sur [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedLocation).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

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

### <a name="remove-location"></a>Supprimer un emplacement

L’exemple suivant montre comment supprimer l’emplacement en appelant [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeAsync_locationIdentifiers__options__callback_) sur [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedLocation).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        Office.context.mailbox.item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>Utiliser `location` l’API

Vous pouvez utiliser `location` l’API pour obtenir et définir l’emplacement d’un rendez-vous.

### <a name="get-the-location"></a>Recherche de l’emplacement

Cette section présente un exemple de code qui obtient et affiche l’emplacement du rendez-vous que compose l’utilisateur.

Pour utiliser `item.location.getAsync`, indiquez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone.  Vous pouvez fournir les arguments nécessaires à la méthode de rappel via le paramètre facultatif `asyncContext`. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie `asyncResult` du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’emplacement sous forme de chaîne à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        getLocation();
    });
}

// Get the location of the item that the user is composing.
function getLocation() {
    item.location.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the location, display it.
                write ('The location is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="set-the-location"></a>Définition de l’emplacement

Cette section présente un exemple de code qui définit l’emplacement du rendez-vous composé par l’utilisateur.

Pour utiliser `item.location.setAsync`, spécifiez une chaîne de 255 caractères maximum dans le paramètre de données. Si vous le souhaitez, vous pouvez fournir une méthode de rappel et tous les arguments de la méthode de rappel dans le paramètre `asyncContext`. Vous devez vérifier l’état, le résultat et tout message d’erreur dans le paramètre `asyncResult` de sortie du rappel. Si l’appel asynchrone aboutit, `setAsync` insère la chaîne d’emplacement spécifiée sous forme de texte brut, en écrasant tous les emplacements existants pour cet élément.

> [!NOTE]
> Vous pouvez définir plusieurs emplacements en utilisant un point-virgule comme séparateur (par exemple, « Salle de conférence A ; Salle de conférence B').

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the location of the item being composed.
        setLocation();
    });
}

// Set the location of the item that the user is composing.
function setLocation() {
    item.location.setAsync(
        'Conference room A',
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the location.
                // Do whatever is appropriate for your scenario,
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

## <a name="see-also"></a>Voir aussi

- [Créer votre premier Outlook de création](../quickstarts/outlook-quickstart.md)
- [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
