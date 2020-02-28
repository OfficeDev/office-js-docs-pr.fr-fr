---
title: Obtenir ou définir l’heure de rendez-vous dans un complément Outlook
description: Découvrez comment obtenir ou définir l’heure de début et de fin d’un rendez-vous dans un complément Outlook.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: d07d461b852e523626946a79a5c9c5e21c95fcdc
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324960"
---
# <a name="get-or-set-the-time-when-composing-an-appointment-in-outlook"></a>Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook

L’API JavaScript pour Office fournit des méthodes asynchrones ([Time. getAsync](/javascript/api/outlook/office.Time#getasync-options--callback-) et [Time. setAsync](/javascript/api/outlook/office.Time#setasync-datetime--options--callback-)) pour obtenir et définir l’heure de début ou de fin d’un rendez-vous composé par l’utilisateur. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, vérifiez que vous avez correctement configuré le manifeste de complément pour Outlook afin d’activer le complément dans les formulaires de composition, comme décrit dans [créer des compléments Outlook pour les formulaires de composition](compose-scenario.md).

Les propriétés [start](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [end](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) sont disponibles pour les rendez-vous dans les formulaires de lecture et de composition. Dans un formulaire de lecture, vous pouvez accéder aux propriétés directement dans l’objet parent, comme dans :

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

Comme pour la plupart des méthodes asynchrones dans l’API JavaScript pour Office, **getAsync** et **setAsync** prennent des paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir [Passage de paramètres facultatifs à des méthodes asynchrones](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) dans [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-the-start-or-end-time"></a>Obtention de l’heure de début ou de fin

Cette section présente un exemple de code qui obtient l’heure de début du rendez-vous que l’utilisateur compose, et affiche cette heure. Vous pouvez utiliser le même code et remplacer la propriété **start** par la propriété **end** pour obtenir l’heure de fin. Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous, comme indiqué ci-dessous.


```XML
<Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>

```

Pour utiliser les éléments **item.start.getAsync** ou **item.end.getAsync**, fournissez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la méthode de rappel via le paramètre facultatif _asyncContext_. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’heure de début comme objet **Date** au format UTC à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).


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


## <a name="set-the-start-or-end-time"></a>Définition de l’heure de début ou de fin

Cette section présente un exemple de code qui définit l’heure de début du rendez-vous ou le message que l’utilisateur compose. Vous pouvez utiliser le même code et remplacer la propriété **start** par la propriété **end** pour définir l’heure de fin. Notez que si le formulaire de composition du rendez-vous contient déjà une heure de début, définir l’heure de début ultérieurement entraînera l’ajustement de l’heure de fin afin de maintenir la durée précédemment définie du rendez-vous. Si le formulaire de composition du rendez-vous contient déjà une heure de fin, définir l’heure de fin ultérieurement entraînera l’ajustement de la durée et de l’heure de fin. Si le rendez-vous a été défini comme un événement d’une journée entière, définir l’heure de début entraînera l’ajustement de l’heure de fin pour la définir à 24 heures plus tard et l’option indiquant qu’il s’agit d’un événement d’une journée entière sera désélectionnée dans le formulaire de composition.

Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste de complément qui active le complément dans un formulaire de composition pour un rendez-vous.

Pour utiliser les éléments **item.start.setAsync** ou **item.end.setAsync**, spécifiez une valeur **Date** au format UTC dans le paramètre _dateTime_. Si vous obtenez une date basée sur une entrée effectuée par l’utilisateur sur le client, vous pouvez utiliser [mailbox.convertToUtcClientTime](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) pour convertir la valeur en objet **Date** au format UTC. Vous pouvez indiquer une méthode de rappel facultative, ainsi que les arguments associés, dans le paramètre _asyncContext_. Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, la méthode **setAsync** insère la chaîne représentant l’heure de début ou de fin spécifiée en tant que texte brut et remplace l’heure de début ou de fin existante pour cet élément.




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


## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)    
- [Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition](item-data.md)   
- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)    
- [Programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-set-or-add-recipients.md)  
- [Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-or-set-the-subject.md)   
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](insert-data-in-the-body.md)   
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-location-of-an-appointment.md)
    
