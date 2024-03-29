---
title: Obtenir ou définir le lieu de rendez-vous dans un complément
description: Découvrez comment obtenir ou définir l’heure d’un rendez-vous à partir d’un complément Outlook.
ms.date: 11/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: d88e2494592d9b261945ecdaf0ca27ae79c73ba8
ms.sourcegitcommit: cae583433e489a3b71418ea270a90db72ad1e838
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/09/2022
ms.locfileid: "68892363"
---
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a>Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook

L’API JavaScript Office fournit des propriétés et des méthodes permettant de gérer l’emplacement d’un rendez-vous que l’utilisateur compose. Actuellement, deux propriétés fournissent l’emplacement d’un rendez-vous :

- [item.location](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) : API de base qui vous permet d’obtenir et de définir l’emplacement.
- [item.enhancedLocation](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) : API améliorée qui vous permet d’obtenir et de définir l’emplacement, et inclut la spécification du [type d’emplacement](/javascript/api/outlook/office.mailboxenums.locationtype). Le type est `LocationType.Custom` si vous définissez l’emplacement à l’aide de `item.location`.

Le tableau suivant répertorie les API d’emplacement et les modes (par exemple, Composer ou Lire) dans lesquels elles sont disponibles.

| API | Modes de rendez-vous applicables |
|---|---|
| [item.location](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-location-member) | Participant/Lecture |
| [item.location.getAsync](/javascript/api/outlook/office.location#outlook-office-location-getasync-member(1)) | Organisateur/Composition |
| [item.location.setAsync](/javascript/api/outlook/office.location#outlook-office-location-setasync-member(1)) | Organisateur/Composition |
| [item.enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) | Organisateur/Composition,<br>Participant/Lecture |
| [item.enhancedLocation.addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) | Organisateur/Composition |
| [item.enhancedLocation.removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) | Organisateur/Composition |

Pour utiliser les méthodes disponibles uniquement pour composer des compléments, configurez le manifeste XML du complément pour activer le complément en mode Organisateur/Composition. Pour plus d’informations, voir [Créer des compléments Outlook pour composer des formulaires](compose-scenario.md) . Les règles d’activation ne sont pas prises en charge dans les compléments qui utilisent un [manifeste Teams pour les compléments Office (préversion).](../develop/json-manifest-overview.md)

## <a name="use-the-enhancedlocation-api"></a>Utiliser l’API `enhancedLocation`

Vous pouvez utiliser l’API `enhancedLocation` pour obtenir et définir l’emplacement d’un rendez-vous. Le champ location prend en charge plusieurs emplacements et, pour chaque emplacement, vous pouvez définir le nom d’affichage, le type et l’adresse e-mail de la salle de conférence (le cas échéant). Consultez [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) pour connaître les types d’emplacements pris en charge.

### <a name="add-location"></a>Ajouter un emplacement

L’exemple suivant montre comment ajouter un emplacement en appelant [addAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-addasync-member(1)) sur [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
let item;
const locations = [
    {
        "id": "Contoso",
        "type": Office.MailboxEnums.LocationType.Custom
    }
];

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add to the location of the item being composed.
        item.enhancedLocation.addAsync(locations);
    });
}
```

### <a name="get-location"></a>Obtenir l’emplacement

L’exemple suivant montre comment obtenir l’emplacement en appelant [getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) sur [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#outlook-office-appointmentread-enhancedlocation-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

> [!NOTE]
> [Les groupes de contacts personnels](https://support.microsoft.com/office/88ff6c60-0a1d-4b54-8c9d-9e1a71bc3023) ajoutés en tant qu’emplacements de rendez-vous ne sont pas retournés par la méthode [enhancedLocation.getAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-getasync-member(1)) .

### <a name="remove-location"></a>Supprimer l’emplacement

L’exemple suivant montre comment supprimer l’emplacement en appelant [removeAsync](/javascript/api/outlook/office.enhancedlocation#outlook-office-enhancedlocation-removeasync-member(1)) sur [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#outlook-office-appointmentcompose-enhancedlocation-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the location of the item being composed.
        item.enhancedLocation.getAsync(callbackFunction);
    });
}

function callbackFunction(asyncResult) {
    asyncResult.value.forEach(function (currentValue) {
        // Remove each location from the item being composed.
        item.enhancedLocation.removeAsync([currentValue.locationIdentifier]);
    });
}
```

## <a name="use-the-location-api"></a>Utiliser l’API `location`

Vous pouvez utiliser l’API `location` pour obtenir et définir l’emplacement d’un rendez-vous.

### <a name="get-the-location"></a>Recherche de l’emplacement

Cette section présente un exemple de code qui obtient et affiche l’emplacement du rendez-vous que compose l’utilisateur.

Pour utiliser `item.location.getAsync`, fournissez une fonction de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la fonction de rappel via le `asyncContext` paramètre facultatif. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre `asyncResult` de sortie du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’emplacement sous forme de chaîne à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

Pour utiliser `item.location.setAsync`, spécifiez une chaîne de 255 caractères maximum dans le paramètre de données. Si vous le souhaitez, vous pouvez fournir une fonction de rappel et tous les arguments pour la fonction de rappel dans le `asyncContext` paramètre . Vous devez vérifier l’état, le résultat et tout message d’erreur dans le `asyncResult` paramètre de sortie du rappel. Si l’appel asynchrone aboutit, `setAsync` insère la chaîne d’emplacement spécifiée sous forme de texte brut, en écrasant tous les emplacements existants pour cet élément.

> [!NOTE]
> Vous pouvez définir plusieurs emplacements à l’aide d’un point-virgule comme séparateur (par exemple, « Salle de conférence A ; Salle de conférence B').

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Check for the DOM to load using the jQuery ready method.
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

- [Créer votre premier complément Outlook](../quickstarts/outlook-quickstart.md)
- [Programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
