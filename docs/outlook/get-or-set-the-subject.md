---
title: Obtenir ou définir l’objet dans un complément Outlook
description: Découvrez comment obtenir ou définir l’objet d’un message ou d’un rendez-vous dans un complément Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 93864aee005af61d9648c39402a843d9105bb021
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325439"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="f93d6-103">Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f93d6-103">Get or set the subject when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="f93d6-104">L’API JavaScript pour Office fournit des méthodes asynchrones ([Subject. getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) et [Subject. setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) pour obtenir et définir l’objet d’un rendez-vous ou d’un message que l’utilisateur compose.</span><span class="sxs-lookup"><span data-stu-id="f93d6-104">The Office JavaScript API provides asynchronous methods ([subject.getAsync](/javascript/api/outlook/office.Subject#getasync-options--callback-) and [subject.setAsync](/javascript/api/outlook/office.Subject#setasync-subject--options--callback-)) to get and set the subject of an appointment or message that the user is composing.</span></span> <span data-ttu-id="f93d6-105">Ces méthodes asynchrones ne sont disponibles que pour les compléments de composition. Pour utiliser ces méthodes, vérifiez que vous avez correctement configuré le manifeste de complément pour Outlook afin d’activer le complément dans les formulaires de composition.</span><span class="sxs-lookup"><span data-stu-id="f93d6-105">These asynchronous methods are available only to compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms.</span></span>

<span data-ttu-id="f93d6-p102">La propriété **subject** est disponible pour un accès en lecture dans les formulaires de lecture et de composition des rendez-vous et des messages. Dans un formulaire de lecture, vous pouvez accéder à la propriété directement à partir de l’objet parent, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="f93d6-p102">The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:</span></span>

```js
item.subject
```

<span data-ttu-id="f93d6-108">Cependant, dans un formulaire de composition, comme l’utilisateur et votre complément peuvent insérer ou modifier l’objet en même temps, vous devez utiliser la méthode asynchrone **getAsync** pour obtenir l’objet, comme indiqué ci-dessous :</span><span class="sxs-lookup"><span data-stu-id="f93d6-108">But in a compose form, because both the user and your add-in can be inserting or changing the subject at the same time, you must use the asynchronous method **getAsync** to get the subject, as shown below:</span></span>

```js
item.subject.getAsync
```

<span data-ttu-id="f93d6-109">La propriété **subject** est disponible pour l’accès en écriture uniquement dans les formulaires de composition, pas dans les formulaires de lecture.</span><span class="sxs-lookup"><span data-stu-id="f93d6-109">The **subject** property is available for write access in only compose forms and not in read forms.</span></span>

<span data-ttu-id="f93d6-110">Comme pour la plupart des méthodes asynchrones dans l’API JavaScript pour Office, **getAsync** et **setAsync** prennent des paramètres d’entrée facultatifs.</span><span class="sxs-lookup"><span data-stu-id="f93d6-110">As with most asynchronous methods in the Office JavaScript API, **getAsync** and **setAsync** take optional input parameters.</span></span> <span data-ttu-id="f93d6-111">Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir la section « Passage de paramètres facultatifs à des méthodes asynchrones » dans la rubrique [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="f93d6-111">For more information about specifying these optional input parameters, see "Passing optional parameters to asynchronous methods" in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-the-subject"></a><span data-ttu-id="f93d6-112">Obtention de l’objet</span><span class="sxs-lookup"><span data-stu-id="f93d6-112">Get the subject</span></span>

<span data-ttu-id="f93d6-p104">Cette section présente un exemple de code qui obtient l’objet du rendez-vous ou du message que l’utilisateur compose, et affiche l’objet. Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="f93d6-p104">This section shows a code sample that gets the subject of the appointment or message that the user is composing, and displays the subject. This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

<span data-ttu-id="f93d6-p105">Pour utiliser **item.subject.getAsync**, fournissez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la méthode de rappel via le paramètre facultatif _asyncContext_. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’objet sous forme de chaîne de texte brut à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).</span><span class="sxs-lookup"><span data-stu-id="f93d6-p105">To use **item.subject.getAsync**, provide a callback method that checks for the status and result of the asynchronous call. You can provide any necessary arguments to the callback method through the  _asyncContext_ optional parameter. You can obtain status, results and any error using the output parameter _asyncResult_ of the callback. If the asynchronous call is successful, you can get the subject as a plain text string using the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property.</span></span>


```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get the subject of the item being composed.
        getSubject();
    });
}

// Get the subject of the item that the user is composing.
function getSubject() {
    item.subject.getAsync(
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the subject, display it.
                write ('The subject is: ' + asyncResult.value);
            }
        });
}

// Write to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-the-subject"></a><span data-ttu-id="f93d6-119">Définition de l’objet</span><span class="sxs-lookup"><span data-stu-id="f93d6-119">Set the subject</span></span>


<span data-ttu-id="f93d6-p106">Cette section présente un exemple de code qui définit l’objet du rendez-vous ou du message que l’utilisateur compose. Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message.</span><span class="sxs-lookup"><span data-stu-id="f93d6-p106">This section shows a code sample that sets the subject of the appointment or message that the user is composing. Similar to the previous example, this code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message.</span></span>

<span data-ttu-id="f93d6-p107">Pour utiliser **item.subject.setAsync**, indiquez une chaîne de 255 caractères maximum dans le paramètre de données. Vous pouvez éventuellement fournir une méthode de rappel et tous les arguments pour la méthode de rappel dans le paramètre _asyncContext_. Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, **setAsync** insère la chaîne d’objet spécifiée sous forme de texte brut, en écrasant tous les objets existants pour cet élément.</span><span class="sxs-lookup"><span data-stu-id="f93d6-p107">To use **item.subject.setAsync**, specify a string of up to 255 characters in the data parameter. Optionally, you can provide a callback method and any arguments for the callback method in the  _asyncContext_ parameter. You should check the status, result and any error message in the _asyncResult_ output parameter of the callback. If the asynchronous call is successful, **setAsync** inserts the specified subject string as plain text, overwriting any existing subject for that item.</span></span>

```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    var today = new Date();
    var subject;

    // Customize the subject with today's date.
    subject = 'Summary for ' + today.toLocaleDateString();

    item.subject.setAsync(
        subject,
        { asyncContext: { var1: 1, var2: 2 } },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully set the subject.
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


## <a name="see-also"></a><span data-ttu-id="f93d6-126">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f93d6-126">See also</span></span>

- [<span data-ttu-id="f93d6-127">Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f93d6-127">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)   
- [<span data-ttu-id="f93d6-128">Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition</span><span class="sxs-lookup"><span data-stu-id="f93d6-128">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="f93d6-129">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="f93d6-129">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="f93d6-130">Programmation asynchrone dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="f93d6-130">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="f93d6-131">Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f93d6-131">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="f93d6-132">Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f93d6-132">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)   
- [<span data-ttu-id="f93d6-133">Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f93d6-133">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="f93d6-134">Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f93d6-134">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
