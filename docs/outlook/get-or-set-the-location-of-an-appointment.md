---
title: Obtenir ou définir le lieu de rendez-vous dans un complément
description: Découvrez comment obtenir ou définir l’heure d’un rendez-vous à partir d’un complément Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 79cf5ebe029d2b95b1501b6f9066a2c8f9013ef3
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609182"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook

L’API JavaScript pour Office fournit des propriétés et des méthodes permettant de gérer l’emplacement d’un rendez-vous que l’utilisateur compose. Actuellement, il existe deux propriétés qui fournissent l’emplacement d’un rendez-vous :

- [Item. Location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): interface API de base qui vous permet d’obtenir et de définir l’emplacement.
- [Item. enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API améliorée qui vous permet d’obtenir et de définir l’emplacement et inclut la spécification du [type d’emplacement](/javascript/api/outlook/office.mailboxenums.locationtype). Le type est `LocationType.Custom` si vous définissez l’emplacement à l’aide du `item.location` .

Le tableau suivant répertorie les API d’emplacement et les modes (par exemple, composition ou lecture) où elles sont disponibles.

| API | Modes de rendez-vous applicables |
|---|---|
| [Item. Location](/javascript/api/outlook/office.appointmentread#location) | Participant/lecture |
| [Item. Location. getAsync](/javascript/api/outlook/office.location#getasync-options--callback-) | Organisateur/composition |
| [item.location.setAsync](/javascript/api/outlook/office.location#setasync-location--options--callback-) | Organisateur/composition |
| [Item. enhancedLocation. getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | Organisateur/composition,<br>Participant/lecture |
| [Item. enhancedLocation. addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | Organisateur/composition |
| [Item. enhancedLocation. removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | Organisateur/composition |

Pour utiliser les méthodes qui sont disponibles uniquement pour les compléments de composition, configurez le manifeste du complément pour activer le complément en mode organisateur/composition. Pour plus d’informations, consultez la rubrique [créer des compléments Outlook pour les formulaires de composition](compose-scenario.md) .

## <a name="use-the-enhancedlocation-api"></a>Utiliser l' `enhancedLocation` API

Vous pouvez utiliser l' `enhancedLocation` API pour obtenir et définir l’emplacement d’un rendez-vous. Le champ emplacement prend en charge plusieurs emplacements et, pour chaque emplacement, vous pouvez définir le nom complet, le type et l’adresse de messagerie de la salle de conférence (le cas échéant). Voir [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) pour les types d’emplacement pris en charge.

### <a name="add-location"></a>Ajouter un emplacement

L’exemple suivant montre comment ajouter un emplacement en appelant [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) sur [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).

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

### <a name="get-location"></a>Obtenir l’emplacement

L’exemple suivant montre comment obtenir l’emplacement en appelant [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) sur [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).

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

### <a name="remove-location"></a>Supprimer l’emplacement

L’exemple suivant montre comment supprimer l’emplacement en appelant [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) sur [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).

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

## <a name="use-the-location-api"></a>Utiliser l' `location` API

Vous pouvez utiliser l' `location` API pour obtenir et définir l’emplacement d’un rendez-vous.

### <a name="get-the-location"></a>Recherche de l’emplacement

Cette section présente un exemple de code qui obtient et affiche l’emplacement du rendez-vous que compose l’utilisateur.

Pour utiliser `item.location.getAsync`, indiquez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone.  Vous pouvez fournir les arguments nécessaires à la méthode de rappel via le paramètre facultatif `asyncContext`. Vous pouvez obtenir l’État, les résultats et toute erreur à l’aide du paramètre `asyncResult` de sortie du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’emplacement sous forme de chaîne à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).

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

Pour utiliser `item.location.setAsync`, spécifiez une chaîne de 255 caractères maximum dans le paramètre de données. Si vous le souhaitez, vous pouvez fournir une méthode de rappel et tous les arguments de la méthode de rappel dans le paramètre `asyncContext`. Vous devez vérifier l’État, le résultat et tous les messages d’erreur dans le `asyncResult` paramètre de sortie du rappel. Si l’appel asynchrone aboutit, `setAsync` insère la chaîne d’emplacement spécifiée sous forme de texte brut, en écrasant tous les emplacements existants pour cet élément.

> [!NOTE]
> Vous pouvez définir plusieurs emplacements à l’aide d’un point-virgule comme séparateur (par exemple, «salle de conférence A ; Salle de conférence B').

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

- [Création de votre premier complément Outlook](../quickstarts/outlook-quickstart.md)
- [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
