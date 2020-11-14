---
title: Ajouter et supprimer des pièces jointes dans un complément Outlook
description: Vous pouvez utiliser différentes API de pièces jointes pour gérer les fichiers ou les éléments Outlook associés à l’élément composé par l’utilisateur.
ms.date: 11/11/2020
localization_priority: Normal
ms.openlocfilehash: 6f146b3efc3234313191d93af05d9c0d35111829
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071703"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a><span data-ttu-id="f00a4-103">Gérer les pièces jointes d’un élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="f00a4-103">Manage an item's attachments in a compose form in Outlook</span></span>

<span data-ttu-id="f00a4-104">L’API JavaScript pour Office fournit plusieurs API que vous pouvez utiliser pour gérer les pièces jointes d’un élément lors de la composition de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="f00a4-104">The Office JavaScript API provides several APIs you can use to manage an item's attachments when the user is composing.</span></span>

## <a name="attach-a-file-or-outlook-item"></a><span data-ttu-id="f00a4-105">Joindre un fichier ou un élément Outlook</span><span class="sxs-lookup"><span data-stu-id="f00a4-105">Attach a file or Outlook item</span></span>

<span data-ttu-id="f00a4-106">Vous pouvez attacher un fichier ou un élément Outlook à un formulaire de composition à l’aide de la méthode appropriée pour le type de pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f00a4-106">You can attach a file or Outlook item to a compose form by using the method that's appropriate for the type of attachment.</span></span>

- <span data-ttu-id="f00a4-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): joindre un fichier</span><span class="sxs-lookup"><span data-stu-id="f00a4-107">[addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file</span></span>
- <span data-ttu-id="f00a4-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): joindre un fichier à l’aide de sa chaîne base64</span><span class="sxs-lookup"><span data-stu-id="f00a4-108">[addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach a file using its base64 string</span></span>
- <span data-ttu-id="f00a4-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): attacher un élément Outlook</span><span class="sxs-lookup"><span data-stu-id="f00a4-109">[addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): Attach an Outlook item</span></span>

<span data-ttu-id="f00a4-110">Il s’agit de méthodes asynchrones, ce qui signifie que l’exécution peut être exécutée sans attendre la fin de l’action.</span><span class="sxs-lookup"><span data-stu-id="f00a4-110">These are asynchronous methods, which means execution can go on without waiting for the action to complete.</span></span> <span data-ttu-id="f00a4-111">En fonction de l’emplacement d’origine et de la taille de la pièce jointe ajoutée, l’appel asynchrone peut prendre un certain temps.</span><span class="sxs-lookup"><span data-stu-id="f00a4-111">Depending on the original location and size of the attachment being added, the asynchronous call may take a while to complete.</span></span>

<span data-ttu-id="f00a4-112">S’il existe des tâches qui dépendent de l’action à effectuer, vous devez les réaliser dans une méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f00a4-112">If there are tasks that depend on the action to complete, you should carry out those tasks in a callback method.</span></span> <span data-ttu-id="f00a4-113">Cette méthode de rappel est facultative et est appelée lorsque le téléchargement de la pièce jointe est terminé.</span><span class="sxs-lookup"><span data-stu-id="f00a4-113">This callback method is optional and is invoked when the attachment upload has completed.</span></span> <span data-ttu-id="f00a4-114">La méthode de rappel utilise un objet [AsyncResult](/javascript/api/office/office.asyncresult) comme paramètre de sortie qui indique les statuts, erreurs et valeurs renvoyés par l’ajout de la pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f00a4-114">The callback method takes an [AsyncResult](/javascript/api/office/office.asyncresult) object as an output parameter that provides any status, error, and returned value from adding the attachment.</span></span> <span data-ttu-id="f00a4-115">Si le rappel requiert des paramètres supplémentaires, vous pouvez les spécifier dans le paramètre facultatif `options.asyncContext`.</span><span class="sxs-lookup"><span data-stu-id="f00a4-115">If the callback requires any extra parameters, you can specify them in the optional `options.asyncContext` parameter.</span></span> <span data-ttu-id="f00a4-116">L’élément `options.asyncContext` peut appartenir à n’importe quel type prévu par votre méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="f00a4-116">`options.asyncContext` can be of any type that your callback method expects.</span></span>

<span data-ttu-id="f00a4-p103">Par exemple, vous pouvez définir `options.asyncContext` comme un objet JSON qui contient au moins une paire clé-valeur. Vous pouvez trouver plus d’exemples sur le passage de paramètres facultatifs à des méthodes asynchrones dans la plateforme des Compléments Office dans [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). L’exemple suivant montre comment utiliser le paramètre asyncContext`asyncContext` pour passer 2 arguments à une méthode de rappel :</span><span class="sxs-lookup"><span data-stu-id="f00a4-p103">For example, you can define `options.asyncContext` as a JSON object that contains one or more key-value pairs. You can find more examples about passing optional parameters to asynchronous methods in the Office Add-ins platform in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-to-asynchronous-methods). The following example shows how to use the `asyncContext` parameter to pass 2 arguments to a callback method:</span></span>

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

<span data-ttu-id="f00a4-p104">Vous pouvez vérifier la réussite ou l’échec d’un appel de méthode asynchrone dans la méthode de rappel à l’aide des propriétés `status` et `error` de l’objet `AsyncResult`. Si l’ajout de pièce jointe aboutit, vous pouvez utiliser la propriété `AsyncResult.value` pour obtenir l’ID de la pièce jointe. Il s’agit d’un nombre entier que vous pouvez ensuite utiliser pour supprimer la pièce jointe.</span><span class="sxs-lookup"><span data-stu-id="f00a4-p104">You can check for success or error of an asynchronous method call in the callback method using the `status` and `error` properties of the `AsyncResult` object. If the attaching completes successfully, you can use the `AsyncResult.value` property to get the attachment ID. The attachment ID is an integer which you can subsequently use to remove the attachment.</span></span>

> [!NOTE]
> <span data-ttu-id="f00a4-122">Nous vous recommandons vivement de ne supprimer une pièce jointe à l’aide de son ID que dans le cas où le même complément a ajouté cette pièce jointe au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="f00a4-122">As a best practice, you should use the attachment ID to remove an attachment only if the same add-in has added that attachment in the same session.</span></span> <span data-ttu-id="f00a4-123">Dans Outlook sur le Web et les appareils mobiles, l’ID de pièce jointe est valide uniquement au sein de la même session.</span><span class="sxs-lookup"><span data-stu-id="f00a4-123">In Outlook on the web and mobile devices, the attachment ID is valid only within the same session.</span></span> <span data-ttu-id="f00a4-124">Une session est terminée lorsque l’utilisateur ferme le complément, ou si l’utilisateur commence la composition dans un formulaire incorporé, avant de fermer ce formulaire pour continuer dans une fenêtre distincte.</span><span class="sxs-lookup"><span data-stu-id="f00a4-124">A session is over when the user closes the add-in, or if the user starts composing in an inline form and subsequently pops out the inline form to continue in a separate window.</span></span>

### <a name="attach-a-file"></a><span data-ttu-id="f00a4-125">Joindre un fichier</span><span class="sxs-lookup"><span data-stu-id="f00a4-125">Attach a file</span></span>

<span data-ttu-id="f00a4-126">Vous pouvez joindre un fichier à un message ou un rendez-vous dans un formulaire de composition à l’aide de la `addFileAttachmentAsync` méthode et en spécifiant l’URI du fichier.</span><span class="sxs-lookup"><span data-stu-id="f00a4-126">You can attach a file to a message or appointment in a compose form by using the `addFileAttachmentAsync` method and specifying the URI of the file.</span></span> <span data-ttu-id="f00a4-127">Vous pouvez également utiliser la `addFileAttachmentFromBase64Async` méthode, mais spécifier la chaîne base64 comme entrée.</span><span class="sxs-lookup"><span data-stu-id="f00a4-127">You can also use the `addFileAttachmentFromBase64Async` method but specify the base64 string as input.</span></span> <span data-ttu-id="f00a4-128">Si le fichier est protégé, vous pouvez inclure une identité appropriée ou un jeton d’authentification comme paramètre de chaîne de requête d’URI.</span><span class="sxs-lookup"><span data-stu-id="f00a4-128">If the file is protected, you can include an appropriate identity or authentication token as a URI query string parameter.</span></span> <span data-ttu-id="f00a4-129">Exchange effectuera un appel à l’URI pour obtenir la pièce jointe, et le service web qui protège le fichier devra utiliser le jeton comme moyen d’authentification.</span><span class="sxs-lookup"><span data-stu-id="f00a4-129">Exchange will make a call to the URI to get the attachment, and the web service which protects the file will need to use the token as a means of authentication.</span></span>

<span data-ttu-id="f00a4-p107">L’exemple JavaScript suivant est un complément de composition qui joint un fichier, picture.png, au message ou au rendez-vous en cours de composition à partir d’un serveur web. La méthode de rappel prend `asyncResult` comme paramètre, vérifie le statut du résultat et obtient l’ID de pièce jointe si la méthode a abouti.</span><span class="sxs-lookup"><span data-stu-id="f00a4-p107">The following JavaScript example is a compose add-in that attaches a file, picture.png, from a web server to the message or appointment being composed. The callback method takes `asyncResult` as a parameter, checks for the result status, and gets the attachment ID if the method succeeds.</span></span>

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback method is invoked and gets the attachment ID.
        // You can optionally pass any object that you would  
        // access in the callback method as an argument to  
        // the asyncContext parameter.
        Office.context.mailbox.item.addFileAttachmentAsync(
            `https://webserver/picture.png`,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                }
                else {
                    // Get the ID of the attached file.
                    var attachmentID = asyncResult.value;
                    write('ID of added attachment: ' + attachmentID);
                }
            });
    });
}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message;
}
```

### <a name="attach-an-outlook-item"></a><span data-ttu-id="f00a4-132">Joindre un élément Outlook</span><span class="sxs-lookup"><span data-stu-id="f00a4-132">Attach an Outlook item</span></span>

<span data-ttu-id="f00a4-p108">Vous pouvez joindre un élément Outlook (par exemple, un élément de messagerie, de calendrier ou de contact) à un message ou à un rendez-vous dans un formulaire de composition en précisant l’ID des services web Exchange (EWS) de l’élément et en utilisant la méthode `addItemAttachmentAsync`. Vous pouvez obtenir l’ID EWS d’un élément de messagerie, de calendrier, de contact ou de tâche dans la boîte aux lettres de l’utilisateur en utilisant la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) et en accédant à l’opération EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). La propriété [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) fournit également l’ID EWS d’un élément existant dans un formulaire de lecture.</span><span class="sxs-lookup"><span data-stu-id="f00a4-p108">You can attach an Outlook item (for example, email, calendar, or contact item) to a message or appointment in a compose form by specifying the Exchange Web Services (EWS) ID of the item and using the `addItemAttachmentAsync` method. You can get the EWS ID of an email, calendar, contact or task item in the user's mailbox by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method and accessing the EWS operation [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). The [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property also provides the EWS ID of an existing item in a read form.</span></span>

<span data-ttu-id="f00a4-136">La fonction JavaScript suivante, `addItemAttachment` , étend le premier exemple ci-dessus et ajoute un élément en tant que pièce jointe à un message électronique ou un rendez-vous composé.</span><span class="sxs-lookup"><span data-stu-id="f00a4-136">The following JavaScript function, `addItemAttachment`, extends the first example above, and adds an item as an attachment to the email or appointment that is being composed.</span></span> <span data-ttu-id="f00a4-137">La fonction prend comme argument l’ID EWS de l’élément qui doit être joint.</span><span class="sxs-lookup"><span data-stu-id="f00a4-137">The function takes as an argument the EWS ID of the item that is to be attached.</span></span> <span data-ttu-id="f00a4-138">Si l’attachement réussit, il obtient l’ID de pièce jointe pour un traitement supplémentaire, y compris la suppression de cette pièce jointe dans la même session.</span><span class="sxs-lookup"><span data-stu-id="f00a4-138">If attaching succeeds, it gets the attachment ID for further processing, including removing that attachment in the same session.</span></span>

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback method is invoked. Here, the callback
    // method uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> <span data-ttu-id="f00a4-139">Vous pouvez utiliser un complément de composition pour joindre une instance d’un rendez-vous périodique dans Outlook sur le Web ou des appareils mobiles.</span><span class="sxs-lookup"><span data-stu-id="f00a4-139">You can use a compose add-in to attach an instance of a recurring appointment in Outlook on the web or mobile devices.</span></span> <span data-ttu-id="f00a4-140">Cependant, dans le client riche Outlook de prise en charge, la tentative d’attachement d’une instance entraîne l’attachement d’une série périodique (rendez-vous principal).</span><span class="sxs-lookup"><span data-stu-id="f00a4-140">However, in a supporting Outlook rich client, attempting to attach an instance would result in attaching the recurring series (the master appointment).</span></span>

## <a name="get-attachments"></a><span data-ttu-id="f00a4-141">Obtention de pièces jointes</span><span class="sxs-lookup"><span data-stu-id="f00a4-141">Get attachments</span></span>

<span data-ttu-id="f00a4-142">Les API permettant d’obtenir des pièces jointes en mode composition sont disponibles à partir de l' [ensemble de conditions requises 1,8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="f00a4-142">APIs to get attachments in compose mode are available from [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

- [<span data-ttu-id="f00a4-143">getAttachmentsAsync</span><span class="sxs-lookup"><span data-stu-id="f00a4-143">getAttachmentsAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [<span data-ttu-id="f00a4-144">getAttachmentContentAsync</span><span class="sxs-lookup"><span data-stu-id="f00a4-144">getAttachmentContentAsync</span></span>](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

<span data-ttu-id="f00a4-145">Vous pouvez utiliser la méthode [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) pour obtenir les pièces jointes du message ou du rendez-vous composé.</span><span class="sxs-lookup"><span data-stu-id="f00a4-145">You can use the [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method to get the attachments of the message or appointment being composed.</span></span>

<span data-ttu-id="f00a4-146">Pour obtenir le contenu d’une pièce jointe, vous pouvez utiliser la méthode [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="f00a4-146">To get an attachment's content, you can use the [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="f00a4-147">Les formats pris en charge sont répertoriés dans l’énumération [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .</span><span class="sxs-lookup"><span data-stu-id="f00a4-147">The supported formats are listed in the [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) enum.</span></span>

<span data-ttu-id="f00a4-148">Vous devez fournir une méthode de rappel pour vérifier l’État et toute erreur à l’aide de l' `AsyncResult` objet de paramètre de sortie.</span><span class="sxs-lookup"><span data-stu-id="f00a4-148">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="f00a4-149">Vous pouvez également transmettre des paramètres supplémentaires à la méthode de rappel à l’aide du paramètre Optional `asyncContext` .</span><span class="sxs-lookup"><span data-stu-id="f00a4-149">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="f00a4-150">L’exemple JavaScript suivant obtient les pièces jointes et vous permet de configurer des opérations de gestion distinctes pour chaque format de pièce jointe pris en charge.</span><span class="sxs-lookup"><span data-stu-id="f00a4-150">The following JavaScript example gets the attachments and allows you to set up distinct handling for each supported attachment format.</span></span>

```js
var item = Office.context.mailbox.item;
var options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (i = 0 ; i < result.value.length ; i++) {
      result.asyncContext.currentItem.getAttachmentContentAsync(result.value[i].id, handleAttachmentsCallback);
    }
  }
}

function handleAttachmentsCallback(result) {
  // Parse string to be a url, an .eml file, a base64-encoded string, or an .icalendar file.
  switch (result.value.format) {
    case Office.MailboxEnums.AttachmentContentFormat.Base64:
      // Handle file attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Eml:
      // Handle email item attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.ICalendar:
      // Handle .icalender attachment.
      break;
    case Office.MailboxEnums.AttachmentContentFormat.Url:
      // Handle cloud attachment.
      break;
    default:
      // Handle attachment formats that are not supported.
  }
}
```

## <a name="remove-an-attachment"></a><span data-ttu-id="f00a4-151">Supprimer une pièce jointe</span><span class="sxs-lookup"><span data-stu-id="f00a4-151">Remove an attachment</span></span>

<span data-ttu-id="f00a4-152">Vous pouvez supprimer une pièce jointe de fichier ou d’élément d’un élément de rendez-vous ou de message dans un formulaire de composition en indiquant l’ID de pièce jointe correspondant et en utilisant la méthode [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods).</span><span class="sxs-lookup"><span data-stu-id="f00a4-152">You can remove a file or item attachment from a message or appointment item in a compose form by specifying the corresponding attachment ID and using the [removeAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="f00a4-153">Vous devez uniquement supprimer les pièces jointes que le même complément a ajouté au cours de la même session.</span><span class="sxs-lookup"><span data-stu-id="f00a4-153">You should only remove attachments that the same add-in has added in the same session.</span></span> <span data-ttu-id="f00a4-154">Semblable aux `addFileAttachmentAsync` méthodes et `addItemAttachmentAsync` , `removeAttachmentAsync` est une méthode asynchrone.</span><span class="sxs-lookup"><span data-stu-id="f00a4-154">Similar to the `addFileAttachmentAsync` and `addItemAttachmentAsync` methods, `removeAttachmentAsync` is an asynchronous method.</span></span> <span data-ttu-id="f00a4-155">Vous devez fournir une méthode de rappel pour vérifier l’État et toute erreur à l’aide de l' `AsyncResult` objet de paramètre de sortie.</span><span class="sxs-lookup"><span data-stu-id="f00a4-155">You should provide a callback method to check for the status and any error by using the `AsyncResult` output parameter object.</span></span> <span data-ttu-id="f00a4-156">Vous pouvez également transmettre des paramètres supplémentaires à la méthode de rappel à l’aide du paramètre Optional `asyncContext` .</span><span class="sxs-lookup"><span data-stu-id="f00a4-156">You can also pass any additional parameters to the callback method by using the optional `asyncContext` parameter.</span></span>

<span data-ttu-id="f00a4-157">La fonction JavaScript suivante, `removeAttachment` continue d’étendre les exemples ci-dessus et supprime la pièce jointe spécifiée de la messagerie ou du rendez-vous composé.</span><span class="sxs-lookup"><span data-stu-id="f00a4-157">The following JavaScript function, `removeAttachment`, continues to extend the examples above, and removes the specified attachment from the email or appointment that is being composed.</span></span> <span data-ttu-id="f00a4-158">La fonction prend comme argument l’ID de la pièce jointe à supprimer.</span><span class="sxs-lookup"><span data-stu-id="f00a4-158">The function takes as an argument the ID of the attachment to be removed.</span></span> <span data-ttu-id="f00a4-159">Vous pouvez obtenir l’ID d’une pièce jointe après un appel de méthode réussi, `addFileAttachmentAsync` `addFileAttachmentFromBase64Async` ou `addItemAttachmentAsync` , et la stocker pour un `removeAttachmentAsync` appel de méthode ultérieur.</span><span class="sxs-lookup"><span data-stu-id="f00a4-159">You can obtain the ID of an attachment after a successful `addFileAttachmentAsync`, `addFileAttachmentFromBase64Async`, or `addItemAttachmentAsync` method call, and store it for a subsequent `removeAttachmentAsync` method call.</span></span>

```js
// Removes the specified attachment from the composed item.
// ID is the Exchange identifier of the attachment to be
// removed.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the
    // callback method is invoked. Here, the callback
    // method uses an asyncResult parameter and gets
    // the ID of the removed attachment if the removal
    // succeeds.
    // You can optionally pass any object you wish to
    // access in the callback method as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.removeAttachmentAsync(
        attachmentId,
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                write(asyncResult.error.message);
            } else {
                write('Removed attachment with the ID: ' + asyncResult.value);
            }
        });
}
```

## <a name="see-also"></a><span data-ttu-id="f00a4-160">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f00a4-160">See also</span></span>

- [<span data-ttu-id="f00a4-161">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="f00a4-161">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)
- [<span data-ttu-id="f00a4-162">Programmation asynchrone dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="f00a4-162">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)
