---
title: Obtenir ou modifier des destinataires dans un complément Outlook
description: Découvrez comment obtenir, définir ou ajouter des destinataires d’un message ou un rendez-vous dans un complément Outlook.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: d6e69b3adc8ddc9f5606e3ec522c56a621eb3664
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609125"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a>Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook


L’API JavaScript pour Office fournit des méthodes asynchrones ([Recipients. getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients. setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)ou [Recipients. addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) pour obtenir, définir ou ajouter respectivement des destinataires dans un formulaire de composition d’un rendez-vous ou d’un message. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, vérifiez que vous avez correctement configuré le manifeste de complément pour Outlook afin d’activer le complément dans les formulaires de composition, comme décrit dans [créer des compléments Outlook pour les formulaires de composition](compose-scenario.md).

Certaines des propriétés qui représentent les destinataires dans un rendez-vous ou un message sont disponibles pour l’accès en lecture dans un formulaire de composition et de lecture. Ces propriétés sont [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour les rendez-vous et [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour les messages. 

Dans un formulaire de lecture, vous pouvez accéder à la propriété directement à partir de l’objet parent, comme :

```js
item.cc
```

Toutefois, dans un formulaire de composition, étant donné que l’utilisateur et votre complément peuvent insérer ou modifier un destinataire en même temps, vous devez utiliser la méthode asynchrone `getAsync` pour obtenir ces propriétés, comme dans l’exemple suivant :


```js
item.cc.getAsync
```

Ces propriétés sont disponibles pour l’accès en écriture uniquement dans les formulaires de composition, pas dans les formulaires de lecture.

Comme pour la plupart des méthodes asynchrones dans l’API JavaScript pour Office,, `getAsync` `setAsync` et `addAsync` prennent des paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir [Passage de paramètres facultatifs à des méthodes asynchrones](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) dans [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="get-recipients"></a>Pour obtenir les destinataires


Cette section présente un exemple de code qui obtient les destinataires d’un rendez-vous ou d’un message dont la composition est en cours et affiche les adresses de messagerie des destinataires. L’exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous.


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Dans l’API JavaScript pour Office, étant donné que les propriétés qui représentent les destinataires d’un rendez-vous ( **optionalAttendees** et **requiredAttendees**) sont différentes de celles d’un message ([BCC](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **CC**et **to**), vous devez d’abord utiliser la propriété [Item. ItemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour identifier si l’élément en cours de composition est un rendez-vous ou un message. En mode composition, toutes les propriétés de rendez-vous et de messages sont des objets [destinataires](/javascript/api/outlook/office.Recipients) , de sorte que vous pouvez appliquer la méthode asynchrone, `Recipients.getAsync` pour obtenir les destinataires correspondants.

Pour utiliser `getAsync` une méthode de rappel, vous pouvez vérifier l’État, les résultats et toute erreur renvoyée par l' `getAsync` appel asynchrone. Vous pouvez fournir des arguments à la méthode de rappel à l’aide du paramètre facultatif  _asyncContext_. La méthode de rappel renvoie un paramètre de sortie  _asyncResult_. Vous pouvez utiliser les `status` `error` Propriétés et de l’objet du paramètre [asyncResult](/javascript/api/office/office.asyncresult) pour vérifier l’État et les messages d’erreur de l’appel asynchrone, et la `value` propriété pour obtenir les destinataires réels. Les destinataires sont représentés dans un tableau d’objets [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).

Notez que, étant donné que la `getAsync` méthode est asynchrone, si des actions ultérieures dépendent de l’obtention réussie des destinataires, vous devez organiser votre code pour démarrer ces actions uniquement dans la méthode de rappel correspondante lorsque l’appel asynchrone s’est correctement terminé.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Get all the recipients of the composed item.
        getAllRecipients();
    });
}

// Get the email addresses of all the recipients of the composed item.
function getAllRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;
    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method getAsync to get each type of recipients
    // of the composed item. Each time, this example passes an anonymous 
    // callback function that doesn't take any parameters.
    toRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get to-recipients of the item completed.
            // Display the email addresses of the to-recipients. 
            write ('To-recipients of the item:');
            displayAddresses(asyncResult);
        }    
    }); // End getAsync for to-recipients.

    // Get any cc-recipients.
    ccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get cc-recipients of the item completed.
            // Display the email addresses of the cc-recipients.
            write ('Cc-recipients of the item:');
            displayAddresses(asyncResult);
        }
    }); // End getAsync for cc-recipients.

    // If the item has the bcc field, i.e., item is message,
    // get any bcc-recipients.
    if (bccRecipients) {
        bccRecipients.getAsync(function (asyncResult) {
        if (asyncResult.status == Office.AsyncResultStatus.Failed){
            write(asyncResult.error.message);
        }
        else {
            // Async call to get bcc-recipients of the item completed.
            // Display the email addresses of the bcc-recipients.
            write ('Bcc-recipients of the item:');
            displayAddresses(asyncResult);
        }
                        
        }); // End getAsync for bcc-recipients.
     }
}

// Recipients are in an array of EmailAddressDetails
// objects passed in asyncResult.value.
function displayAddresses (asyncResult) {
    for (var i=0; i<asyncResult.value.length; i++)
        write (asyncResult.value[i].emailAddress);
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="set-recipients"></a>Définir les destinataires


Cette section présente un exemple de code qui définit les destinataires du rendez-vous ou du message que l’utilisateur compose. Le fait de définir des destinataires remplace tous les destinataires existants. Comme dans l’exemple précédent relatif à l’obtention des destinataires dans un formulaire de composition, cet exemple suppose que le complément est activé dans les formulaires de composition pour les rendez-vous et les messages. Cet exemple vérifie d’abord si l’élément composé est un rendez-vous ou un message, pour appliquer la méthode asynchrone, `Recipients.setAsync` , sur les propriétés appropriées qui représentent les destinataires du rendez-vous ou du message.

Lors de `setAsync` l’appel, fournissez un tableau comme argument d’entrée pour le paramètre _Recipients_ , dans l’un des formats suivants :


- Un tableau de chaînes représentant des adresses SMTP.
    
- Un tableau de dictionnaires, chacun contenant un nom d’affichage et une adresse de messagerie, comme indiqué dans l’exemple de code suivant.
    
- Tableau d' `EmailAddressDetails` objets, semblable à celui renvoyé par la `getAsync` méthode.
    
Vous pouvez éventuellement fournir une méthode de rappel comme argument d’entrée à la `setAsync` méthode, afin de vous assurer que tout code qui dépend de la définition réussie des destinataires ne s’exécute que lorsque cela se produit. Vous pouvez également fournir des arguments à la méthode de rappel à l’aide du paramètre facultatif _asyncContext_. Si vous utilisez une méthode de rappel, vous pouvez accéder à un paramètre de sortie _asyncResult_ et utiliser les propriétés **Status** et **Error** de l' `AsyncResult` objet Parameter pour vérifier l’État et les messages d’erreur de l’appel asynchrone.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set recipients of the composed item.
        setRecipients();
    });
}

// Set the display name and email addresses of the recipients of 
// the composed item.
function setRecipients() {
    // Local objects to point to recipients of either
    // the appointment or message that is being composed.
    // bccRecipients applies to only messages, not appointments.
    var toRecipients, ccRecipients, bccRecipients;

    // Verify if the composed item is an appointment or message.
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        toRecipients = item.requiredAttendees;
        ccRecipients = item.optionalAttendees;
    }
    else {
        toRecipients = item.to;
        ccRecipients = item.cc;
        bccRecipients = item.bcc;
    }
    
    // Use asynchronous method setAsync to set each type of recipients
    // of the composed item. Each time, this example passes a set of
    // names and email addresses to set, and an anonymous 
    // callback function that doesn't take any parameters. 
    toRecipients.setAsync(
        [{
            "displayName":"Graham Durkin", 
            "emailAddress":"graham@contoso.com"
         },
         {
            "displayName" : "Donnie Weinberg",
            "emailAddress" : "donnie@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set to-recipients of the item completed.

            }    
    }); // End to setAsync.


    // Set any cc-recipients.
    ccRecipients.setAsync(
        [{
             "displayName":"Perry Horning", 
             "emailAddress":"perry@contoso.com"
         },
         {
             "displayName" : "Guy Montenegro",
             "emailAddress" : "guy@contoso.com"
         }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to set cc-recipients of the item completed.
            }
    }); // End cc setAsync.


    // If the item has the bcc field, i.e., item is message,
    // set bcc-recipients.
    if (bccRecipients) {
        bccRecipients.setAsync(
            [{
                 "displayName":"Lewis Cate", 
                 "emailAddress":"lewis@contoso.com"
             },
             {
                 "displayName" : "Francisco Stitt",
                 "emailAddress" : "francisco@contoso.com"
             }],
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Async call to set bcc-recipients of the item completed.
                    // Do whatever appropriate for your scenario.
                }
        }); // End bcc setAsync.
    }
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}

```


## <a name="add-recipients"></a>Ajouter des destinataires

Si vous ne souhaitez pas remplacer les destinataires existants dans un rendez-vous ou un message, au lieu d’utiliser `Recipients.setAsync` , vous pouvez utiliser la `Recipients.addAsync` méthode asynchrone pour ajouter des destinataires. `addAsync`fonctionne de la même manière que `setAsync` dans la mesure où il nécessite un argument d’entrée de _destinataires_ . Vous pouvez éventuellement fournir une méthode de rappel et tous les arguments pour le rappel à l’aide du paramètre asyncContext. Vous pouvez ensuite vérifier l’État, le résultat et toute erreur de l' `addAsync` appel asynchrone en utilisant le paramètre de sortie _asyncResult_ de la méthode de rappel. L’exemple suivant vérifie que l’élément en cours de composition est un rendez-vous et y ajoute deux participants obligatoires.


```js
// Add specified recipients as required attendees of
// the composed appointment. 
function addAttendees() {
    if (item.itemType == Office.MailboxEnums.ItemType.Appointment) {
        item.requiredAttendees.addAsync(
        [{
            "displayName":"Kristie Jensen", 
            "emailAddress":"kristie@contoso.com"
         },
         {
            "displayName" : "Pansy Valenzuela",
            "emailAddress" : "pansy@contoso.com"
          }],
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Async call to add attendees completed.
                // Do whatever appropriate for your scenario.
            }
        }); // End addAsync.
    }
}
```


## <a name="see-also"></a>Voir aussi

- [Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook](get-and-set-item-data-in-a-compose-form.md)
- [Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition](item-data.md)
- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)
- [Programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
- [Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook](get-or-set-the-subject.md)
- [Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook](insert-data-in-the-body.md)
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-location-of-an-appointment.md)
- [Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-time-of-an-appointment.md)
    
