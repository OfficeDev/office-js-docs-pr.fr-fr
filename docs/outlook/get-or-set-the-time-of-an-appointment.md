---
title: Obtenir ou définir l’heure de rendez-vous dans un complément Outlook
description: Découvrez comment obtenir ou définir l’heure de début et de fin d’un rendez-vous dans un complément Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: af4ec04c8f7af865c826a036b6670c0aec7341b4
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166115"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a><span data-ttu-id="ca256-103">Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="ca256-103">Get or set the time when composing an appointment in Outlook</span></span>

<span data-ttu-id="ca256-p101">L’interface de l’API JavaScript pour Office fournit des méthodes asynchrones ([Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) et [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) pour obtenir et définir l’heure de début ou de fin d’un rendez-vous composé par l’utilisateur. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, assurez-vous que vous avez correctement configuré le manifeste du complément pour Outlook afin d’activer le complément dans des formulaires de composition, comme décrit dans la rubrique [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="ca256-p101">The JavaScript API for Office provides asynchronous methods ([Time.getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) and [Time.setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) to get and set the start or end time of an appointment that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="ca256-p102">Les propriétés [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) sont disponibles pour les rendez-vous dans les formulaires de lecture et de composition. Dans un formulaire de lecture, vous pouvez accéder aux propriétés directement dans l’objet parent, comme dans :</span><span class="sxs-lookup"><span data-stu-id="ca256-p102">The [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:</span></span>

```js
item.start
```

<span data-ttu-id="ca256-108">et dans :</span><span class="sxs-lookup"><span data-stu-id="ca256-108">and in:</span></span>

```js
item.end
```

<span data-ttu-id="ca256-109">Cependant, dans un formulaire de composition, comme l’utilisateur et votre complément peuvent insérer ou modifier l’heure en même temps, vous devez utiliser la méthode asynchrone **getAsync** pour obtenir l’heure de début ou de fin, comme indiqué ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="ca256-109">But in a compose form, because both the user and your add-in can be inserting or changing the time at the same time, you must use the asynchronous method **getAsync** to get the start or end time, as shown below:</span></span>

```js
item.start.getAsync
```

<span data-ttu-id="ca256-110">et :</span><span class="sxs-lookup"><span data-stu-id="ca256-110">and:</span></span>

```js
item.end.getAsync
```

<span data-ttu-id="ca256-p103">Comme avec la plupart des méthodes asynchrones dans l’interface API JavaScript pour Office, les méthodes **getAsync** et **setAsync** admettent des paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, consultez la section [Passage de paramètres facultatifs à des méthodes asynchrones](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) dans la rubrique [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="ca256-p103">As with most asynchronous methods in the JavaScript API for Office, **getAsync** and **setAsync** take optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-start-or-end-time"></a><span data-ttu-id="ca256-113">Obtention de l’heure de début ou de fin</span><span class="sxs-lookup"><span data-stu-id="ca256-113">Get the start or end time</span></span>

<span data-ttu-id="ca256-p104">Cette section présente un exemple de code qui obtient l’heure de début du rendez-vous que l’utilisateur compose, et affiche cette heure. Vous pouvez utiliser le même code et remplacer la propriété **start** par la propriété **end** pour obtenir l’heure de fin. Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous, comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="ca256-p104">This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.</span></span>


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

<span data-ttu-id="ca256-p105">Pour utiliser les éléments **item.start.getAsync** ou **item.end.getAsync**, fournissez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la méthode de rappel via le paramètre facultatif _asyncContext_. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’heure de début comme objet **Date** au format UTC à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="ca256-p105">To use **item.start.getAsync** or **item.end.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the start time as a **Date** object in UTC format using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the start time of the item being composed.
        getStartTime();
    });
}

// Get the start time of the item that the user is composing.
function getStartTime() {
    item.start.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the start time, display it, first in UTC and 
                // then convert the Date object to local time and display that.
                write ('The start time in UTC is: ' + asyncResult.value.toString());
                write ('The start time in local time is: ' + asyncResult.value.toLocaleString());
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-start-or-end-time"></a><span data-ttu-id="ca256-121">Définition de l’heure de début ou de fin</span><span class="sxs-lookup"><span data-stu-id="ca256-121">Set the start or end time</span></span>

<span data-ttu-id="ca256-p106">Cette section présente un exemple de code qui définit l’heure de début du rendez-vous ou le message que l’utilisateur compose. Vous pouvez utiliser le même code et remplacer la propriété **start** par la propriété **end** pour définir l’heure de fin. Notez que si le formulaire de composition du rendez-vous contient déjà une heure de début, définir l’heure de début ultérieurement entraînera l’ajustement de l’heure de fin afin de maintenir la durée précédemment définie du rendez-vous. Si le formulaire de composition du rendez-vous contient déjà une heure de fin, définir l’heure de fin ultérieurement entraînera l’ajustement de la durée et de l’heure de fin. Si le rendez-vous a été défini comme un événement d’une journée entière, définir l’heure de début entraînera l’ajustement de l’heure de fin pour la définir à 24 heures plus tard et l’option indiquant qu’il s’agit d’un événement d’une journée entière sera désélectionnée dans le formulaire de composition.</span><span class="sxs-lookup"><span data-stu-id="ca256-p106">This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.</span></span>

<span data-ttu-id="ca256-127">Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste de complément qui active le complément dans un formulaire de composition pour un rendez-vous.</span><span class="sxs-lookup"><span data-stu-id="ca256-127">Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment.</span></span>

<span data-ttu-id="ca256-p107">Pour utiliser les éléments **item.start.setAsync** ou **item.end.setAsync**, spécifiez une valeur **Date** au format UTC dans le paramètre _dateTime_. Si vous obtenez une date basée sur une entrée effectuée par l’utilisateur sur le client, vous pouvez utiliser [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour convertir la valeur en objet **Date** au format UTC. Vous pouvez indiquer une méthode de rappel facultative, ainsi que les arguments associés, dans le paramètre _asyncContext_. Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, la méthode **setAsync** insère la chaîne représentant l’heure de début ou de fin spécifiée en tant que texte brut et remplace l’heure de début ou de fin existante pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="ca256-p107">To use **item.start.setAsync** or **item.end.setAsync**, specify a **Date** value in UTC in the _dateTime_ parameter. If you get a date based on an input by the user on the client, you can use [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) to convert the value to a **Date** object in UTC. You can provide an optional callback method and any arguments for the callback method in the _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified start or end time string as plain text, overwriting any existing start or end time for that item.</span></span>




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    var startDate = new Date("September 27, 2012 12:30:00");
    
    item.start.setAsync(
        startDate,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the start time.
                // Do whatever appropriate for your scenario
                // using the arguments var1 and var2 as applicable.
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="see-also"></a><span data-ttu-id="ca256-133">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="ca256-133">See also</span></span>

- [<span data-ttu-id="ca256-134">Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="ca256-134">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="ca256-135">Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition</span><span class="sxs-lookup"><span data-stu-id="ca256-135">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)   
- [<span data-ttu-id="ca256-136">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="ca256-136">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="ca256-137">Programmation asynchrone dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="ca256-137">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="ca256-138">Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="ca256-138">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="ca256-139">Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="ca256-139">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)   
- [<span data-ttu-id="ca256-140">Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="ca256-140">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="ca256-141">Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="ca256-141">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
    
