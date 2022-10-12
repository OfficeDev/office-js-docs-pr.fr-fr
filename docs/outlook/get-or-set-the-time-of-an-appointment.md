---
title: Obtenir ou définir l’heure de rendez-vous dans un complément Outlook
description: Découvrez comment obtenir ou définir l’heure de début et de fin d’un rendez-vous dans un complément Outlook.
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: c7aa40fda15c613aca869af8b277d4deb6fbf833
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541232"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook

L’API JavaScript Office fournit des méthodes asynchrones ([Time.getAsync](/javascript/api/outlook/office.time#outlook-office-time-getasync-member(1)) et [Time.setAsync](/javascript/api/outlook/office.time#outlook-office-time-setasync-member(1))) pour obtenir et définir l’heure de début ou de fin d’un rendez-vous que l’utilisateur compose. Ces méthodes asynchrones sont disponibles uniquement pour composer des compléments. Pour utiliser ces méthodes, assurez-vous d’avoir correctement configuré le manifeste XML du complément pour qu’Outlook active le complément dans les formulaires de composition, comme décrit dans [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md). Les règles d’activation ne sont pas prises en charge dans les compléments qui utilisent un [manifeste Teams pour les compléments Office (préversion).](../develop/json-manifest-overview.md)

The [start](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) and [end](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) properties are available for appointments in both compose and read forms. In a read form, you can access the properties directly from the parent object, as in:

```js
item.start
```

et dans :

```js
item.end
```

Cependant, dans un formulaire de composition, comme l’utilisateur et votre complément peuvent insérer ou modifier l’heure en même temps, vous devez utiliser la méthode asynchrone **getAsync** pour obtenir l’heure de début ou de fin, comme indiqué ci-dessous :

```js
item.start.getAsync
```

et :

```js
item.end.getAsync
```

Comme avec la plupart des méthodes asynchrones dans l’API JavaScript Office, **getAsync** et **setAsync prennent des paramètres d’entrée** facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir [Passage de paramètres facultatifs à des méthodes asynchrones](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-inline) dans [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-the-start-or-end-time"></a>Obtention de l’heure de début ou de fin

This section shows a code sample that gets the start time of the appointment that the user is composing and displays the time. You can use the same code and replace the **start** property by the **end** property to get the end time. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment, as shown below.

```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
```

Pour utiliser **item.start.getAsync** ou **item.end.getAsync**, fournissez une fonction de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la fonction de rappel via le paramètre facultatif  _asyncContext_ . Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie  _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’heure de début comme objet **Date** au format UTC à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

## <a name="set-the-start-or-end-time"></a>Définition de l’heure de début ou de fin

This section shows a code sample that sets the start time of the appointment or message that the user is composing. You can use the same code and replace the **start** property by the **end** property to set the end time. Note that if the appointment compose form already has an existing start time, setting the start time subsequently will adjust the end time to maintain any previous duration for the appointment. If the appointment compose form already has an existing end time, setting the end time subsequently will adjust both the duration and end time. If the appointment has been set as an all-day event, setting the start time will adjust the end time to 24 hours later, and uncheck the UI for the all-day event in the compose form.

Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste de complément qui active le complément dans un formulaire de composition pour un rendez-vous.

Pour utiliser **item.start.setAsync** ou **item.end.setAsync**, spécifiez une valeur **Date** en UTC dans le paramètre _dateTime_ . Si vous obtenez une date basée sur une entrée effectuée par l’utilisateur sur le client, vous pouvez utiliser [mailbox.convertToUtcClientTime](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) pour convertir la valeur en objet **Date** au format UTC. Vous pouvez fournir une fonction de rappel facultative et tous les arguments de la fonction de rappel dans le paramètre _asyncContext_ . Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie  _asyncResult_ du rappel. Si l’appel asynchrone aboutit, **setAsync** insère la chaîne représentant l’heure de début ou de fin spécifiée en tant que texte brut et remplace l’heure de début ou de fin existante pour cet élément.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the start time of the item being composed.
        setStartTime();
    });
}

// Set the start time of the item that the user is composing.
function setStartTime() {
    const startDate = new Date("September 27, 2012 12:30:00");
    
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

## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition](item-data.md)
- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)
- [Programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-set-or-add-recipients.md)  
- [Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-or-set-the-subject.md)
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](insert-data-in-the-body.md)
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-location-of-an-appointment.md)
