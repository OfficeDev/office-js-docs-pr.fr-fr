---
title: Obtenir ou définir l’objet dans un complément Outlook
description: Découvrez comment obtenir ou définir l’objet d’un message ou d’un rendez-vous dans un complément Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 048aa079adf3fda5d5f4a85bfcadd3b671ce865a
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937509"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook

L’API JavaScript Office fournit des méthodes asynchrones ([subject.getAsync](/javascript/api/outlook/office.Subject#getAsync_options__callback_) et [subject.setAsync](/javascript/api/outlook/office.subject#setAsync_subject__options__callback_)) pour obtenir et définir l’objet d’un rendez-vous ou d’un message que l’utilisateur compose. Ces méthodes asynchrones sont disponibles uniquement pour composer des add-ins. Pour utiliser ces méthodes, assurez-vous que vous avez correctement installé le manifeste du Outlook pour activer le module dans les formulaires de composition.

La propriété **subject** est disponible pour un accès en lecture dans les formulaires de lecture et de composition des rendez-vous et des messages. Dans un formulaire de lecture, vous pouvez accéder à la propriété directement à partir de l’objet parent, comme dans l’exemple suivant :

```js
item.subject
```

Cependant, dans un formulaire de composition, comme l’utilisateur et votre complément peuvent insérer ou modifier l’objet en même temps, vous devez utiliser la méthode asynchrone **getAsync** pour obtenir l’objet, comme indiqué ci-dessous :

```js
item.subject.getAsync
```

La propriété **subject** est disponible pour l’accès en écriture uniquement dans les formulaires de composition, pas dans les formulaires de lecture.

Comme avec la plupart des méthodes asynchrones dans l’API JavaScript Office, **getAsync** et **setAsync** prennent des paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir la section « Passage de paramètres facultatifs à des méthodes asynchrones » dans la rubrique [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-the-subject"></a>Obtention de l’objet

Cette section présente un exemple de code qui obtient l’objet du rendez-vous ou du message que l’utilisateur compose, et affiche l’objet. Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```

Pour utiliser **item.subject.getAsync**, fournissez une méthode de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la méthode de rappel via le paramètre facultatif _asyncContext_. Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’objet sous forme de chaîne de texte brut à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value).


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


## <a name="set-the-subject"></a>Définition de l’objet


Cette section présente un exemple de code qui définit l’objet du rendez-vous ou du message que l’utilisateur compose. Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message.

Pour utiliser **item.subject.setAsync**, indiquez une chaîne de 255 caractères maximum dans le paramètre de données. Vous pouvez éventuellement fournir une méthode de rappel et tous les arguments pour la méthode de rappel dans le paramètre _asyncContext_. Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie _asyncResult_ du rappel. Si l’appel asynchrone aboutit, **setAsync** insère la chaîne d’objet spécifiée sous forme de texte brut, en écrasant tous les objets existants pour cet élément.

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


## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)   
- [Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition](item-data.md)    
- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)    
- [Programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-set-or-add-recipients.md)  
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](insert-data-in-the-body.md)   
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-time-of-an-appointment.md)
    
