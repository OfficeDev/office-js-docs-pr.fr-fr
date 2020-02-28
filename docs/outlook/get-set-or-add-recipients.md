---
title: Obtenir ou modifier des destinataires dans un complément Outlook
description: Découvrez comment obtenir, définir ou ajouter des destinataires d’un message ou un rendez-vous dans un complément Outlook.
ms.date: 12/10/2019
localization_priority: Normal
ms.openlocfilehash: 396f425f639c0d7043154ccfe1ddea16a236f993
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325432"
---
# <a name="get-set-or-add-recipients-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="4bf88-103">Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="4bf88-103">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>


<span data-ttu-id="4bf88-104">L’API JavaScript pour Office fournit des méthodes asynchrones ([Recipients. getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients. setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-)ou [Recipients. addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) pour obtenir, définir ou ajouter respectivement des destinataires dans un formulaire de composition d’un rendez-vous ou d’un message.</span><span class="sxs-lookup"><span data-stu-id="4bf88-104">The Office JavaScript API provides asynchronous methods ([Recipients.getAsync](/javascript/api/outlook/office.Recipients#getasync-options--callback-), [Recipients.setAsync](/javascript/api/outlook/office.Recipients#setasync-recipients--options--callback-), or [Recipients.addAsync](/javascript/api/outlook/office.Recipients#addasync-recipients--options--callback-)) to respectively get, set, or add recipients in a compose form of an appointment or message.</span></span> <span data-ttu-id="4bf88-105">Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, vérifiez que vous avez correctement configuré le manifeste de complément pour Outlook afin d’activer le complément dans les formulaires de composition, comme décrit dans [créer des compléments Outlook pour les formulaires de composition](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="4bf88-105">These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately for Outlook to activate the add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="4bf88-p102">Certaines des propriétés qui représentent les destinataires dans un rendez-vous ou un message sont disponibles pour l’accès en lecture dans un formulaire de composition et de lecture. Ces propriétés sont [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour les rendez-vous et [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) et [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour les messages. </span><span class="sxs-lookup"><span data-stu-id="4bf88-p102">Some of the properties that represent recipients in an appointment or message are available for read access in a compose form and in a read form. These properties include  [optionalAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) and [requiredAttendees](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for appointments, and [cc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), and  [to](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) for messages.</span></span>

<span data-ttu-id="4bf88-108">Dans un formulaire de lecture, vous pouvez accéder à la propriété directement à partir de l’objet parent, comme :</span><span class="sxs-lookup"><span data-stu-id="4bf88-108">In a read form, you can access the property directly from the parent object, such as:</span></span>

```js
item.cc
```

<span data-ttu-id="4bf88-109">Toutefois, dans un formulaire de composition, étant donné que l’utilisateur et votre complément peuvent insérer ou modifier un destinataire en même temps, vous devez utiliser la méthode `getAsync` asynchrone pour obtenir ces propriétés, comme dans l’exemple suivant :</span><span class="sxs-lookup"><span data-stu-id="4bf88-109">But in a compose form, because both the user and your add-in can be inserting or changing a recipient at the same time, you must use the asynchronous method `getAsync` to get these properties, as in the following example:</span></span>


```js
item.cc.getAsync
```

<span data-ttu-id="4bf88-110">Ces propriétés sont disponibles pour l’accès en écriture uniquement dans les formulaires de composition, pas dans les formulaires de lecture.</span><span class="sxs-lookup"><span data-stu-id="4bf88-110">These properties are available for write access in only compose forms and not read forms.</span></span>

<span data-ttu-id="4bf88-111">Comme pour la plupart des méthodes asynchrones dans l’API JavaScript `getAsync`pour `setAsync`Office, `addAsync` , et prennent des paramètres d’entrée facultatifs.</span><span class="sxs-lookup"><span data-stu-id="4bf88-111">As with most asynchronous methods in the JavaScript API for Office, `getAsync`, `setAsync`, and `addAsync` take optional input parameters.</span></span> <span data-ttu-id="4bf88-112">Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir [Passage de paramètres facultatifs à des méthodes asynchrones](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) dans [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="4bf88-112">For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="get-recipients"></a><span data-ttu-id="4bf88-113">Pour obtenir les destinataires</span><span class="sxs-lookup"><span data-stu-id="4bf88-113">Get recipients</span></span>


<span data-ttu-id="4bf88-p104">Cette section présente un exemple de code qui obtient les destinataires d’un rendez-vous ou d’un message dont la composition est en cours et affiche les adresses de messagerie des destinataires. L’exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="4bf88-p104">This section shows a code sample that gets the recipients of the appointment or message that is being composed, and displays the email addresses of the recipients. The code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>


```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

<span data-ttu-id="4bf88-116">Dans l’API JavaScript pour Office, étant donné que les propriétés qui représentent les destinataires d’un rendez-vous ( **optionalAttendees** et **requiredAttendees**) sont différentes de celles d’un message ([BCC](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **CC**et **to**), vous devez d’abord utiliser la propriété [Item. ItemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) pour identifier si l’élément en cours de composition est un rendez-vous ou un message.</span><span class="sxs-lookup"><span data-stu-id="4bf88-116">In the Office JavaScript API, because the properties that represent the recipients of an appointment ( **optionalAttendees** and **requiredAttendees**) are different from those of a message ([bcc](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), **cc**, and **to**), you should first use the [item.itemType](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property to identify whether the item being composed is an appointment or message.</span></span> <span data-ttu-id="4bf88-117">En mode composition, toutes les propriétés de rendez-vous et de messages sont des objets [destinataires](/javascript/api/outlook/office.Recipients) , de sorte que vous `Recipients.getAsync`pouvez appliquer la méthode asynchrone, pour obtenir les destinataires correspondants.</span><span class="sxs-lookup"><span data-stu-id="4bf88-117">In compose mode, all these properties of appointments and messages are [Recipients](/javascript/api/outlook/office.Recipients) objects, so you can then apply the asynchronous method, `Recipients.getAsync`, to get the corresponding recipients.</span></span>

<span data-ttu-id="4bf88-118">Pour utiliser `getAsync` une méthode de rappel, vous pouvez vérifier l’État, les résultats et toute erreur renvoyée par `getAsync` l’appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4bf88-118">To use `getAsync` provide a callback method to check for the status, results, and any error returned by the asynchronous `getAsync` call.</span></span> <span data-ttu-id="4bf88-119">Vous pouvez fournir des arguments à la méthode de rappel à l’aide du paramètre facultatif  _asyncContext_.</span><span class="sxs-lookup"><span data-stu-id="4bf88-119">You can provide any arguments to the callback method using the optional _asyncContext_ parameter.</span></span> <span data-ttu-id="4bf88-120">La méthode de rappel renvoie un paramètre de sortie  _asyncResult_.</span><span class="sxs-lookup"><span data-stu-id="4bf88-120">The callback method returns an _asyncResult_ output parameter.</span></span> <span data-ttu-id="4bf88-121">Vous pouvez utiliser les `status` propriétés `error` et de l’objet du paramètre [asyncResult](/javascript/api/office/office.asyncresult) pour vérifier l’État et les messages d’erreur de l’appel asynchrone, `value` et la propriété pour obtenir les destinataires réels.</span><span class="sxs-lookup"><span data-stu-id="4bf88-121">You can use the `status` and `error` properties of the [AsyncResult](/javascript/api/office/office.asyncresult) parameter object to check for status and any error messages of the asynchronous call, and the `value` property to get the actual recipients.</span></span> <span data-ttu-id="4bf88-122">Les destinataires sont représentés dans un tableau d’objets [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails).</span><span class="sxs-lookup"><span data-stu-id="4bf88-122">Recipients are represented as an array of [EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails) objects.</span></span>

<span data-ttu-id="4bf88-123">Notez que, étant `getAsync` donné que la méthode est asynchrone, si des actions ultérieures dépendent de l’obtention réussie des destinataires, vous devez organiser votre code pour démarrer ces actions uniquement dans la méthode de rappel correspondante lorsque l’appel asynchrone s’est correctement terminé.</span><span class="sxs-lookup"><span data-stu-id="4bf88-123">Note that because the `getAsync` method is asynchronous, if there are subsequent actions that depend on successfully getting the recipients, you should organize your code to start such actions only in the corresponding callback method when the asynchronous call has successfully completed.</span></span>




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


## <a name="set-recipients"></a><span data-ttu-id="4bf88-124">Définir les destinataires</span><span class="sxs-lookup"><span data-stu-id="4bf88-124">Set recipients</span></span>


<span data-ttu-id="4bf88-125">Cette section présente un exemple de code qui définit les destinataires du rendez-vous ou du message que l’utilisateur compose.</span><span class="sxs-lookup"><span data-stu-id="4bf88-125">This section shows a code sample that sets the recipients of the appointment or message that is being composed by the user.</span></span> <span data-ttu-id="4bf88-126">Le fait de définir des destinataires remplace tous les destinataires existants.</span><span class="sxs-lookup"><span data-stu-id="4bf88-126">Setting recipients overwrites any existing recipients.</span></span> <span data-ttu-id="4bf88-127">Comme dans l’exemple précédent relatif à l’obtention des destinataires dans un formulaire de composition, cet exemple suppose que le complément est activé dans les formulaires de composition pour les rendez-vous et les messages.</span><span class="sxs-lookup"><span data-stu-id="4bf88-127">Similar to the previous example that gets recipients in a compose form, this example assumes that the add-in is activated in compose forms for appointments and messages.</span></span> <span data-ttu-id="4bf88-128">Cet exemple vérifie d’abord si l’élément composé est un rendez-vous ou un message, pour appliquer la `Recipients.setAsync`méthode asynchrone,, sur les propriétés appropriées qui représentent les destinataires du rendez-vous ou du message.</span><span class="sxs-lookup"><span data-stu-id="4bf88-128">This example first verifies if the composed item is an appointment or message, so to apply the asynchronous method, `Recipients.setAsync`, on the appropriate properties that represent recipients of the appointment or message.</span></span>

<span data-ttu-id="4bf88-129">Lors de `setAsync`l’appel, fournissez un tableau comme argument d’entrée pour le paramètre _Recipients_ , dans l’un des formats suivants :</span><span class="sxs-lookup"><span data-stu-id="4bf88-129">When calling `setAsync`, provide an array as input argument for the  _recipients_ parameter, in one of the following formats:</span></span>


- <span data-ttu-id="4bf88-130">Un tableau de chaînes représentant des adresses SMTP.</span><span class="sxs-lookup"><span data-stu-id="4bf88-130">An array of strings that are SMTP addresses.</span></span>
    
- <span data-ttu-id="4bf88-131">Un tableau de dictionnaires, chacun contenant un nom d’affichage et une adresse de messagerie, comme indiqué dans l’exemple de code suivant.</span><span class="sxs-lookup"><span data-stu-id="4bf88-131">An array of dictionaries, each containing a display name and email address, as shown in the following code sample.</span></span>
    
- <span data-ttu-id="4bf88-132">Tableau d' `EmailAddressDetails` objets, semblable à celui renvoyé par la `getAsync` méthode.</span><span class="sxs-lookup"><span data-stu-id="4bf88-132">An array of `EmailAddressDetails` objects, similar to the one returned by the `getAsync` method.</span></span>
    
<span data-ttu-id="4bf88-133">Vous pouvez éventuellement fournir une méthode de rappel comme argument d’entrée à la `setAsync` méthode, afin de vous assurer que tout code qui dépend de la définition réussie des destinataires ne s’exécute que lorsque cela se produit.</span><span class="sxs-lookup"><span data-stu-id="4bf88-133">You can optionally provide a callback method as an input argument to the `setAsync` method, to make sure any code that depends on successfully setting the recipients would execute only when that happens.</span></span> <span data-ttu-id="4bf88-134">Vous pouvez également fournir des arguments à la méthode de rappel à l’aide du paramètre facultatif _asyncContext_.</span><span class="sxs-lookup"><span data-stu-id="4bf88-134">You can also provide any arguments for the callback method using the optional _asyncContext_ parameter.</span></span> <span data-ttu-id="4bf88-135">Si vous utilisez une méthode de rappel, vous pouvez accéder à un paramètre de sortie _asyncResult_ et utiliser les propriétés **Status** et `AsyncResult` **Error** de l’objet Parameter pour vérifier l’État et les messages d’erreur de l’appel asynchrone.</span><span class="sxs-lookup"><span data-stu-id="4bf88-135">If you use a callback method, you can access an _asyncResult_ output parameter, and use the **status** and **error** properties of the `AsyncResult` parameter object to check for status and any error messages of the asynchronous call.</span></span>




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


## <a name="add-recipients"></a><span data-ttu-id="4bf88-136">Ajouter des destinataires</span><span class="sxs-lookup"><span data-stu-id="4bf88-136">Add recipients</span></span>

<span data-ttu-id="4bf88-137">Si vous ne souhaitez pas remplacer les destinataires existants dans un rendez-vous ou un message, `Recipients.setAsync`au lieu d’utiliser, `Recipients.addAsync` vous pouvez utiliser la méthode asynchrone pour ajouter des destinataires.</span><span class="sxs-lookup"><span data-stu-id="4bf88-137">If you do not want to overwrite any existing recipients in an appointment or message, instead of using `Recipients.setAsync`, you can use the `Recipients.addAsync` asynchronous method to append recipients.</span></span> <span data-ttu-id="4bf88-138">`addAsync`fonctionne de la même `setAsync` manière que dans la mesure où il nécessite un argument d’entrée de _destinataires_ .</span><span class="sxs-lookup"><span data-stu-id="4bf88-138">`addAsync` works similarly as `setAsync` in that it requires a _recipients_ input argument.</span></span> <span data-ttu-id="4bf88-139">Vous pouvez éventuellement fournir une méthode de rappel et tous les arguments pour le rappel à l’aide du paramètre asyncContext.</span><span class="sxs-lookup"><span data-stu-id="4bf88-139">You can optionally provide a callback method, and any arguments for the callback using the asyncContext parameter.</span></span> <span data-ttu-id="4bf88-140">Vous pouvez ensuite vérifier l’État, le résultat et toute erreur de l’appel `addAsync` asynchrone en utilisant le paramètre de sortie _asyncResult_ de la méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="4bf88-140">You can then check the status, result, and any error of the asynchronous `addAsync` call by using the _asyncResult_ output parameter of the callback method.</span></span> <span data-ttu-id="4bf88-141">L’exemple suivant vérifie que l’élément en cours de composition est un rendez-vous et y ajoute deux participants obligatoires.</span><span class="sxs-lookup"><span data-stu-id="4bf88-141">The following example checks if the item being composed is an appointment, and appends two required attendees to the appointment.</span></span>


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


## <a name="see-also"></a><span data-ttu-id="4bf88-142">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="4bf88-142">See also</span></span>

- [<span data-ttu-id="4bf88-143">Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="4bf88-143">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="4bf88-144">Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition</span><span class="sxs-lookup"><span data-stu-id="4bf88-144">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)
- [<span data-ttu-id="4bf88-145">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="4bf88-145">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="4bf88-146">Programmation asynchrone dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="4bf88-146">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
- [<span data-ttu-id="4bf88-147">Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="4bf88-147">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="4bf88-148">Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="4bf88-148">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="4bf88-149">Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="4bf88-149">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="4bf88-150">Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="4bf88-150">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
