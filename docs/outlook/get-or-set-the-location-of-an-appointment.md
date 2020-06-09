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
# <a name="get-or-set-the-location-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="3c62d-103">Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="3c62d-103">Get or set the location when composing an appointment in Outlook</span></span>

<span data-ttu-id="3c62d-104">L’API JavaScript pour Office fournit des propriétés et des méthodes permettant de gérer l’emplacement d’un rendez-vous que l’utilisateur compose.</span><span class="sxs-lookup"><span data-stu-id="3c62d-104">The Office JavaScript API provides properties and methods to manage the location of an appointment that the user is composing.</span></span> <span data-ttu-id="3c62d-105">Actuellement, il existe deux propriétés qui fournissent l’emplacement d’un rendez-vous :</span><span class="sxs-lookup"><span data-stu-id="3c62d-105">Currently, there are two properties that provide an appointment's location:</span></span>

- <span data-ttu-id="3c62d-106">[Item. Location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): interface API de base qui vous permet d’obtenir et de définir l’emplacement.</span><span class="sxs-lookup"><span data-stu-id="3c62d-106">[item.location](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Basic API that allows you to get and set the location.</span></span>
- <span data-ttu-id="3c62d-107">[Item. enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): API améliorée qui vous permet d’obtenir et de définir l’emplacement et inclut la spécification du [type d’emplacement](/javascript/api/outlook/office.mailboxenums.locationtype).</span><span class="sxs-lookup"><span data-stu-id="3c62d-107">[item.enhancedLocation](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties): Enhanced API that allows you to get and set the location, and includes specifying the [location type](/javascript/api/outlook/office.mailboxenums.locationtype).</span></span> <span data-ttu-id="3c62d-108">Le type est `LocationType.Custom` si vous définissez l’emplacement à l’aide du `item.location` .</span><span class="sxs-lookup"><span data-stu-id="3c62d-108">The type is `LocationType.Custom` if you set the location using `item.location`.</span></span>

<span data-ttu-id="3c62d-109">Le tableau suivant répertorie les API d’emplacement et les modes (par exemple, composition ou lecture) où elles sont disponibles.</span><span class="sxs-lookup"><span data-stu-id="3c62d-109">The following table lists the location APIs and the modes (i.e., Compose or Read) where they are available.</span></span>

| <span data-ttu-id="3c62d-110">API</span><span class="sxs-lookup"><span data-stu-id="3c62d-110">API</span></span> | <span data-ttu-id="3c62d-111">Modes de rendez-vous applicables</span><span class="sxs-lookup"><span data-stu-id="3c62d-111">Applicable appointment modes</span></span> |
|---|---|
| [<span data-ttu-id="3c62d-112">Item. Location</span><span class="sxs-lookup"><span data-stu-id="3c62d-112">item.location</span></span>](/javascript/api/outlook/office.appointmentread#location) | <span data-ttu-id="3c62d-113">Participant/lecture</span><span class="sxs-lookup"><span data-stu-id="3c62d-113">Attendee/Read</span></span> |
| [<span data-ttu-id="3c62d-114">Item. Location. getAsync</span><span class="sxs-lookup"><span data-stu-id="3c62d-114">item.location.getAsync</span></span>](/javascript/api/outlook/office.location#getasync-options--callback-) | <span data-ttu-id="3c62d-115">Organisateur/composition</span><span class="sxs-lookup"><span data-stu-id="3c62d-115">Organizer/Compose</span></span> |
| [<span data-ttu-id="3c62d-116">item.location.setAsync</span><span class="sxs-lookup"><span data-stu-id="3c62d-116">item.location.setAsync</span></span>](/javascript/api/outlook/office.location#setasync-location--options--callback-) | <span data-ttu-id="3c62d-117">Organisateur/composition</span><span class="sxs-lookup"><span data-stu-id="3c62d-117">Organizer/Compose</span></span> |
| [<span data-ttu-id="3c62d-118">Item. enhancedLocation. getAsync</span><span class="sxs-lookup"><span data-stu-id="3c62d-118">item.enhancedLocation.getAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) | <span data-ttu-id="3c62d-119">Organisateur/composition,</span><span class="sxs-lookup"><span data-stu-id="3c62d-119">Organizer/Compose,</span></span><br><span data-ttu-id="3c62d-120">Participant/lecture</span><span class="sxs-lookup"><span data-stu-id="3c62d-120">Attendee/Read</span></span> |
| [<span data-ttu-id="3c62d-121">Item. enhancedLocation. addAsync</span><span class="sxs-lookup"><span data-stu-id="3c62d-121">item.enhancedLocation.addAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) | <span data-ttu-id="3c62d-122">Organisateur/composition</span><span class="sxs-lookup"><span data-stu-id="3c62d-122">Organizer/Compose</span></span> |
| [<span data-ttu-id="3c62d-123">Item. enhancedLocation. removeAsync</span><span class="sxs-lookup"><span data-stu-id="3c62d-123">item.enhancedLocation.removeAsync</span></span>](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) | <span data-ttu-id="3c62d-124">Organisateur/composition</span><span class="sxs-lookup"><span data-stu-id="3c62d-124">Organizer/Compose</span></span> |

<span data-ttu-id="3c62d-125">Pour utiliser les méthodes qui sont disponibles uniquement pour les compléments de composition, configurez le manifeste du complément pour activer le complément en mode organisateur/composition.</span><span class="sxs-lookup"><span data-stu-id="3c62d-125">To use the methods that are available only to compose add-ins, configure the add-in manifest to activate the add-in in Organizer/Compose mode.</span></span> <span data-ttu-id="3c62d-126">Pour plus d’informations, consultez la rubrique [créer des compléments Outlook pour les formulaires de composition](compose-scenario.md) .</span><span class="sxs-lookup"><span data-stu-id="3c62d-126">See [Create Outlook add-ins for compose forms](compose-scenario.md) for more details.</span></span>

## <a name="use-the-enhancedlocation-api"></a><span data-ttu-id="3c62d-127">Utiliser l' `enhancedLocation` API</span><span class="sxs-lookup"><span data-stu-id="3c62d-127">Use the `enhancedLocation` API</span></span>

<span data-ttu-id="3c62d-128">Vous pouvez utiliser l' `enhancedLocation` API pour obtenir et définir l’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="3c62d-128">You can use the `enhancedLocation` API to get and set an appointment's location.</span></span> <span data-ttu-id="3c62d-129">Le champ emplacement prend en charge plusieurs emplacements et, pour chaque emplacement, vous pouvez définir le nom complet, le type et l’adresse de messagerie de la salle de conférence (le cas échéant).</span><span class="sxs-lookup"><span data-stu-id="3c62d-129">The location field supports multiple locations and, for each location, you can set the display name, type, and conference room email address (if applicable).</span></span> <span data-ttu-id="3c62d-130">Voir [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) pour les types d’emplacement pris en charge.</span><span class="sxs-lookup"><span data-stu-id="3c62d-130">See [LocationType](/javascript/api/outlook/office.mailboxenums.locationtype) for supported location types.</span></span>

### <a name="add-location"></a><span data-ttu-id="3c62d-131">Ajouter un emplacement</span><span class="sxs-lookup"><span data-stu-id="3c62d-131">Add location</span></span>

<span data-ttu-id="3c62d-132">L’exemple suivant montre comment ajouter un emplacement en appelant [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) sur [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="3c62d-132">The following example shows how to add a location by calling [addAsync](/javascript/api/outlook/office.enhancedlocation#addasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

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

### <a name="get-location"></a><span data-ttu-id="3c62d-133">Obtenir l’emplacement</span><span class="sxs-lookup"><span data-stu-id="3c62d-133">Get location</span></span>

<span data-ttu-id="3c62d-134">L’exemple suivant montre comment obtenir l’emplacement en appelant [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) sur [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="3c62d-134">The following example shows how to get the location by calling [getAsync](/javascript/api/outlook/office.enhancedlocation#getasync-options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread#enhancedlocation).</span></span>

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

### <a name="remove-location"></a><span data-ttu-id="3c62d-135">Supprimer l’emplacement</span><span class="sxs-lookup"><span data-stu-id="3c62d-135">Remove location</span></span>

<span data-ttu-id="3c62d-136">L’exemple suivant montre comment supprimer l’emplacement en appelant [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) sur [Mailbox. Item. enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span><span class="sxs-lookup"><span data-stu-id="3c62d-136">The following example shows how to remove the location by calling [removeAsync](/javascript/api/outlook/office.enhancedlocation#removeasync-locationidentifiers--options--callback-) on [mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentcompose#enhancedlocation).</span></span>

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

## <a name="use-the-location-api"></a><span data-ttu-id="3c62d-137">Utiliser l' `location` API</span><span class="sxs-lookup"><span data-stu-id="3c62d-137">Use the `location` API</span></span>

<span data-ttu-id="3c62d-138">Vous pouvez utiliser l' `location` API pour obtenir et définir l’emplacement d’un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="3c62d-138">You can use the `location` API to get and set an appointment's location.</span></span>

### <a name="get-the-location"></a><span data-ttu-id="3c62d-139">Recherche de l’emplacement</span><span class="sxs-lookup"><span data-stu-id="3c62d-139">Get the location</span></span>

<span data-ttu-id="3c62d-140">Cette section présente un exemple de code qui obtient et affiche l’emplacement du rendez-vous que compose l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3c62d-140">This section shows a code sample that gets the location of the appointment that the user is composing, and displays the location.</span></span>

<span data-ttu-id="3c62d-141">Pour utiliser `item.location.getAsync`, indiquez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone. </span><span class="sxs-lookup"><span data-stu-id="3c62d-141">To use `item.location.getAsync`, provide a callback method that checks for the status and result of the asynchronous call.</span></span> <span data-ttu-id="3c62d-142">Vous pouvez fournir les arguments nécessaires à la méthode de rappel via le paramètre facultatif `asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="3c62d-142">You can provide any necessary arguments to the callback method through the `asyncContext` optional parameter.</span></span> <span data-ttu-id="3c62d-143">Vous pouvez obtenir l’État, les résultats et toute erreur à l’aide du paramètre `asyncResult` de sortie du rappel.</span><span class="sxs-lookup"><span data-stu-id="3c62d-143">You can obtain status, results, and any error using the output parameter `asyncResult` of the callback.</span></span> <span data-ttu-id="3c62d-144">Si l’appel asynchrone aboutit, vous pouvez obtenir l’emplacement sous forme de chaîne à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="3c62d-144">If the asynchronous call is successful, you can get the location as a string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>

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

### <a name="set-the-location"></a><span data-ttu-id="3c62d-145">Définition de l’emplacement</span><span class="sxs-lookup"><span data-stu-id="3c62d-145">Set the location</span></span>

<span data-ttu-id="3c62d-146">Cette section présente un exemple de code qui définit l’emplacement du rendez-vous composé par l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="3c62d-146">This section shows a code sample that sets the location of the appointment that the user is composing.</span></span>

<span data-ttu-id="3c62d-147">Pour utiliser `item.location.setAsync`, spécifiez une chaîne de 255 caractères maximum dans le paramètre de données.</span><span class="sxs-lookup"><span data-stu-id="3c62d-147">To use `item.location.setAsync`, specify a string of up to 255 characters in the data parameter.</span></span> <span data-ttu-id="3c62d-148">Si vous le souhaitez, vous pouvez fournir une méthode de rappel et tous les arguments de la méthode de rappel dans le paramètre `asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="3c62d-148">Optionally, you can provide a callback method and any arguments for the callback method in the `asyncContext` parameter.</span></span> <span data-ttu-id="3c62d-149">Vous devez vérifier l’État, le résultat et tous les messages d’erreur dans le `asyncResult` paramètre de sortie du rappel.</span><span class="sxs-lookup"><span data-stu-id="3c62d-149">You should check the status, result, and any error message in the `asyncResult` output parameter of the callback.</span></span> <span data-ttu-id="3c62d-150">Si l’appel asynchrone aboutit, `setAsync` insère la chaîne d’emplacement spécifiée sous forme de texte brut, en écrasant tous les emplacements existants pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="3c62d-150">If the asynchronous call is successful, `setAsync` inserts the specified location string as plain text, overwriting any existing location for that item.</span></span>

> [!NOTE]
> <span data-ttu-id="3c62d-151">Vous pouvez définir plusieurs emplacements à l’aide d’un point-virgule comme séparateur (par exemple, «salle de conférence A ; Salle de conférence B').</span><span class="sxs-lookup"><span data-stu-id="3c62d-151">You can set multiple locations by using a semi-colon as the separator (e.g., 'Conference room A; Conference room B').</span></span>

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

## <a name="see-also"></a><span data-ttu-id="3c62d-152">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="3c62d-152">See also</span></span>

- [<span data-ttu-id="3c62d-153">Création de votre premier complément Outlook</span><span class="sxs-lookup"><span data-stu-id="3c62d-153">Create your first Outlook add-in</span></span>](../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="3c62d-154">Programmation asynchrone dans des compléments Office</span><span class="sxs-lookup"><span data-stu-id="3c62d-154">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
