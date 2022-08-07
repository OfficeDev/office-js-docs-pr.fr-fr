---
title: Ajouter et supprimer des pièces jointes dans un complément Outlook
description: Utilisez différentes API de pièce jointe pour gérer les fichiers ou les éléments Outlook attachés à l’élément que l’utilisateur compose.
ms.date: 08/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: af3b44814fd11c5e2006dbb921130c15c7535385
ms.sourcegitcommit: 76b8c79cba707c771ae25df57df14b6445f9b8fa
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/07/2022
ms.locfileid: "67274168"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Gérer les pièces jointes d’un élément dans un formulaire de composition dans Outlook

L’API JavaScript Office fournit plusieurs API que vous pouvez utiliser pour gérer les pièces jointes d’un élément lorsque l’utilisateur compose.

## <a name="attach-a-file-or-outlook-item"></a>Attacher un fichier ou un élément Outlook

Vous pouvez attacher un fichier ou un élément Outlook à un formulaire de composition à l’aide de la méthode appropriée pour le type de pièce jointe.

- [addFileAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) : joindre un fichier
- [addFileAttachmentFromBase64Async](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) : joindre un fichier à l’aide de sa chaîne base64
- [addItemAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) : attacher un élément Outlook

Il s’agit de méthodes asynchrones, ce qui signifie que l’exécution peut continuer sans attendre la fin de l’action. Selon l’emplacement d’origine et la taille de la pièce jointe ajoutée, l’appel asynchrone peut prendre un certain temps.

S’il existe des tâches qui dépendent de l’action à effectuer, vous devez effectuer ces tâches dans une fonction de rappel. Cette fonction de rappel est facultative et est appelée lorsque le chargement de la pièce jointe est terminé. La fonction de rappel prend un objet [AsyncResult](/javascript/api/office/office.asyncresult) comme paramètre de sortie qui fournit l’état, l’erreur et la valeur retournée de l’ajout de la pièce jointe. Si le rappel requiert des paramètres supplémentaires, vous pouvez les spécifier dans le paramètre facultatif `options.asyncContext`. `options.asyncContext` peut être de n’importe quel type attendu par votre fonction de rappel.

Par exemple, vous pouvez définir `options.asyncContext` comme un objet JSON qui contient une ou plusieurs paires clé-valeur. Vous trouverez d’autres exemples sur la transmission de paramètres facultatifs à des méthodes asynchrones dans la plateforme de compléments Office dans la [programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods). L’exemple suivant montre comment utiliser le `asyncContext` paramètre pour passer 2 arguments à une fonction de rappel.

```js
const options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Vous pouvez vérifier la réussite ou l’erreur d’un appel de méthode asynchrone dans la fonction de rappel à l’aide des `status` propriétés et `error` des propriétés de l’objet `AsyncResult` . Si l’attachement se termine correctement, vous pouvez utiliser la `AsyncResult.value` propriété pour obtenir l’ID de pièce jointe. Il s’agit d’un nombre entier que vous pouvez ensuite utiliser pour supprimer la pièce jointe.

> [!NOTE]
> L’ID de pièce jointe est valide uniquement dans la même session et il n’est pas garanti qu’il soit mappé à la même pièce jointe entre les sessions. Par exemple, lorsqu’une session est terminée, citons le moment où l’utilisateur ferme le complément, ou si l’utilisateur commence à composer dans un formulaire inline et sort par la suite le formulaire inline pour continuer dans une fenêtre distincte.

### <a name="attach-a-file"></a>Joindre un fichier

Vous pouvez joindre un fichier à un message ou un rendez-vous dans un formulaire de composition à l’aide de la `addFileAttachmentAsync` méthode et en spécifiant l’URI du fichier. Vous pouvez également utiliser la `addFileAttachmentFromBase64Async` méthode, mais spécifier la chaîne base64 comme entrée. Si le fichier est protégé, vous pouvez inclure une identité appropriée ou un jeton d’authentification comme paramètre de chaîne de requête d’URI. Exchange effectuera un appel à l’URI pour obtenir la pièce jointe, et le service web qui protège le fichier devra utiliser le jeton comme moyen d’authentification.

L’exemple JavaScript suivant est un complément de composition qui joint un fichier, picture.png, à partir d’un serveur web au message ou rendez-vous en cours de composition. La fonction de rappel prend `asyncResult` comme paramètre, vérifie l’état du résultat et obtient l’ID de pièce jointe si la méthode réussit.

```js
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Add the specified file attachment to the item
        // being composed.
        // When the attachment finishes uploading, the
        // callback function is invoked and gets the attachment ID.
        // You can optionally pass any object that you would
        // access in the callback function as an argument to
        // the asyncContext parameter.
        Office.context.mailbox.item.addFileAttachmentAsync(
            `https://webserver/picture.png`,
            'picture.png',
            { asyncContext: null },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                } else {
                    // Get the ID of the attached file.
                    const attachmentID = asyncResult.value;
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

Pour ajouter une image base64 inline au corps d’un message ou d’un rendez-vous en cours de composition, vous devez d’abord obtenir le corps de l’élément actif à l’aide de la `Office.context.mailbox.item.body.getAsync` méthode avant d’insérer l’image à l’aide de la `addFileAttachmentFromBase64Async` méthode. Sinon, l’image ne s’affiche pas dans le corps une fois insérée. Pour obtenir des conseils, consultez l’exemple JavaScript suivant, qui ajoute une image base64 inline au début d’un corps d’élément.

```js
const mailItem = Office.context.mailbox.item;
const base64String =
  "iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAnUExURQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAN0S+bUAAAAMdFJOUwAQIDBAUI+fr7/P7yEupu8AAAAJcEhZcwAADsMAAA7DAcdvqGQAAAF8SURBVGhD7dfLdoMwDEVR6Cspzf9/b20QYOthS5Zn0Z2kVdY6O2WULrFYLBaLxd5ur4mDZD14b8ogWS/dtxV+dmx9ysA2QUj9TQRWv5D7HyKwuIW9n0vc8tkpHP0W4BOg3wQ8wtlvA+PC1e8Ao8Ld7wFjQtHvAiNC2e8DdqHqKwCrUPc1gE1AfRVgEXBfB+gF0lcCWoH2tYBOYPpqQCNwfT3QF9i+AegJfN8CtAWhbwJagtS3AbIg9o2AJMh9M5C+SVGBvx6zAfmT0r+Bv8JMwP4kyFPir+cswF5KL3WLv14zAFBCLf56Tw9cparFX4upgaJUtPhrOS1QlY5W+vWTXrGgBFB/b72ev3/0igUdQPppP/nfowfKUUEFcP207y/yxKmgAYQ+PywoAFOfCH3A2MdCFzD3kdADBvq10AGG+pXQBgb7pdAEhvuF0AIc/VtoAK7+JciAs38KIuDugyAC/v4hiMCE/i7IwLRBsh68N2WQjMVisVgs9i5bln8LGScNcCrONQAAAABJRU5ErkJggg==";

// Get the current body of the message or appointment.
mailItem.body.getAsync(Office.CoercionType.Html, (bodyResult) => {
  if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
    // Insert the base64 image to the beginning of the body.
    const options = { isInline: true, asyncContext: bodyResult.value };
    mailItem.addFileAttachmentFromBase64Async(base64String, "sample.png", options, (attachResult) => {
      if (attachResult.status === Office.AsyncResultStatus.Succeeded) {
        let body = attachResult.asyncContext;
        body = body.replace("<p class=MsoNormal>", `<p class=MsoNormal><img src="cid:sample.png">`);
        mailItem.body.setAsync(body, { coercionType: Office.CoercionType.Html }, (setResult) => {
          if (setResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("Inline base64 image added to the body.");
          } else {
            console.log(setResult.error.message);
          }
        });
      } else {
        console.log(attachResult.error.message);
      }
    });
  } else {
    console.log(bodyResult.error.message);
  }
});
```

### <a name="attach-an-outlook-item"></a>Attacher un élément Outlook

Vous pouvez attacher un élément Outlook (par exemple, un e-mail, un calendrier ou un élément de contact) à un message ou un rendez-vous dans un formulaire de composition en spécifiant l’ID EWS (Exchange Web Services) de l’élément et en utilisant la `addItemAttachmentAsync` méthode. Vous pouvez obtenir l’ID EWS d’un e-mail, d’un calendrier, d’un contact ou d’un élément de tâche dans la boîte aux lettres de l’utilisateur à l’aide de la méthode [mailbox.makeEwsRequestAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) et en accédant à l’opération EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). La propriété [item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties) fournit également l’ID EWS d’un élément existant dans un formulaire de lecture.

La fonction JavaScript suivante, `addItemAttachment`étend le premier exemple ci-dessus et ajoute un élément en tant que pièce jointe à l’e-mail ou au rendez-vous en cours de composition. La fonction prend comme argument l’ID EWS de l’élément qui doit être joint. Si l’attachement réussit, il obtient l’ID de pièce jointe pour un traitement ultérieur, y compris la suppression de cette pièce jointe dans la même session.

```js
// Adds the specified item as an attachment to the composed item.
// ID is the EWS ID of the item to be attached.
function addItemAttachment(itemId) {
    // When the attachment finishes uploading, the
    // callback function is invoked. Here, the callback
    // function uses only asyncResult as a parameter,
    // and if the attaching succeeds, gets the attachment ID.
    // You can optionally pass any other object you wish to
    // access in the callback function as an argument to
    // the asyncContext parameter.
    Office.context.mailbox.item.addItemAttachmentAsync(
        itemId,
        'Welcome email',
        { asyncContext: null },
        function (asyncResult) {
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                const attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Vous pouvez utiliser un complément de composition pour attacher une instance d’un rendez-vous périodique dans Outlook sur le web ou sur des appareils mobiles. Toutefois, dans un client de bureau Outlook pris en charge, toute tentative d’attachement d’une instance entraînerait l’attachement de la série périodique (le rendez-vous parent).

## <a name="get-attachments"></a>Obtention de pièces jointes

Les API permettant d’obtenir des pièces jointes en mode composition sont disponibles à partir de [l’ensemble de conditions requises 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8).

- [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)
- [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods)

Vous pouvez utiliser la méthode [getAttachmentsAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) pour obtenir les pièces jointes du message ou du rendez-vous en cours de composition.

Pour obtenir le contenu d’une pièce jointe, vous pouvez utiliser la méthode [getAttachmentContentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Les formats pris en charge sont répertoriés dans l’énumération [AttachmentContentFormat](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat) .

Vous devez fournir une fonction de rappel pour vérifier l’état et toute erreur à l’aide de l’objet `AsyncResult` de paramètre de sortie. Vous pouvez également passer tous les paramètres supplémentaires à la fonction de rappel à l’aide du paramètre facultatif `asyncContext` .

L’exemple JavaScript suivant obtient les pièces jointes et vous permet de configurer une gestion distincte pour chaque format de pièce jointe pris en charge.

```js
const item = Office.context.mailbox.item;
const options = {asyncContext: {currentItem: item}};
item.getAttachmentsAsync(options, callback);

function callback(result) {
  if (result.value.length > 0) {
    for (let i = 0 ; i < result.value.length ; i++) {
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

## <a name="remove-an-attachment"></a>Supprimer une pièce jointe

Vous pouvez supprimer un fichier ou une pièce jointe d’un message ou d’un élément de rendez-vous dans un formulaire de composition en spécifiant l’ID de pièce jointe correspondant lors de l’utilisation de la méthode [removeAttachmentAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) .

> [!IMPORTANT]
> Si vous utilisez l’ensemble de conditions requises 1.7 ou version antérieure, vous devez supprimer uniquement les pièces jointes ajoutées par le même complément dans la même session.

Similaire à la `addFileAttachmentAsync`méthode , `addItemAttachmentAsync`et `getAttachmentsAsync` aux méthodes, `removeAttachmentAsync` est une méthode asynchrone. Vous devez fournir une fonction de rappel pour vérifier l’état et toute erreur à l’aide de l’objet `AsyncResult` de paramètre de sortie. Vous pouvez également passer tous les paramètres supplémentaires à la fonction de rappel à l’aide du paramètre facultatif `asyncContext` .

La fonction JavaScript suivante, `removeAttachment`continue d’étendre les exemples ci-dessus et supprime la pièce jointe spécifiée de l’e-mail ou du rendez-vous en cours de composition. La fonction prend comme argument l’ID de la pièce jointe à supprimer. Vous pouvez obtenir l’ID d’une pièce jointe après un `addFileAttachmentAsync``addFileAttachmentFromBase64Async`appel de méthode réussi, ou `addItemAttachmentAsync` un appel de méthode, et l’utiliser dans un appel de méthode suivant`removeAttachmentAsync`. Vous pouvez également appeler `getAttachmentsAsync` (introduit dans l’ensemble de conditions requises 1.8) pour obtenir les pièces jointes et leurs ID pour cette session de complément.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback function is invoked.
    // Here, the callback function uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback function as an argument to the asyncContext parameter.
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

## <a name="see-also"></a>Voir aussi

- [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md)
- [Programmation asynchrone dans les compléments Office](../develop/asynchronous-programming-in-office-add-ins.md)
