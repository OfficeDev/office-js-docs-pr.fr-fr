---
title: Ajouter et supprimer des pièces jointes dans un complément Outlook
description: Vous pouvez utiliser différentes API de pièce jointe pour gérer les fichiers ou Outlook éléments joints à l’élément que l’utilisateur compose.
ms.date: 07/08/2021
ms.localizationpriority: medium
ms.openlocfilehash: 1df37568d52ae76caaec3e65286fac5544b577c3
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153216"
---
# <a name="manage-an-items-attachments-in-a-compose-form-in-outlook"></a>Gérer les pièces jointes d’un élément dans un formulaire de composition Outlook

L Office API JavaScript fournit plusieurs API que vous pouvez utiliser pour gérer les pièces jointes d’un élément lorsque l’utilisateur compose.

## <a name="attach-a-file-or-outlook-item"></a>Joindre un fichier ou un Outlook’élément

Vous pouvez joindre un fichier ou un Outlook à un formulaire de composition à l’aide de la méthode appropriée pour le type de pièce jointe.

- [addFileAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): joindre un fichier
- [addFileAttachmentFromBase64Async](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): joindre un fichier à l’aide de sa chaîne base64
- [addItemAttachmentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods): joindre un Outlook’élément

Il s’agit de méthodes asynchrones, ce qui signifie que l’exécution peut continuer sans attendre la fin de l’action. Selon l’emplacement d’origine et la taille de la pièce jointe ajoutée, l’appel asynchrone peut prendre un certain temps.

S’il existe des tâches qui dépendent de l’action à effectuer, vous devez les réaliser dans une méthode de rappel. Cette méthode de rappel est facultative et est invoquée lorsque le chargement de la pièce jointe est terminé. La méthode de rappel utilise un objet [AsyncResult](/javascript/api/office/office.asyncresult) comme paramètre de sortie qui indique les statuts, erreurs et valeurs renvoyés par l’ajout de la pièce jointe. Si le rappel requiert des paramètres supplémentaires, vous pouvez les spécifier dans le paramètre facultatif `options.asyncContext`. L’élément `options.asyncContext` peut appartenir à n’importe quel type prévu par votre méthode de rappel.

Par exemple, vous pouvez définir en `options.asyncContext` tant qu’objet JSON qui contient une ou plusieurs paires clé-valeur. Vous trouverez d’autres exemples sur la transmission de paramètres facultatifs à des méthodes asynchrones dans la plateforme de Office Add-ins en [programmation asynchrone dans Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md#pass-optional-parameters-to-asynchronous-methods). L’exemple suivant montre comment utiliser le paramètre `asyncContext` pour passer 2 arguments à une méthode de rappel.

```js
var options = { asyncContext: { var1: 1, var2: 2}};

Office.context.mailbox.item.addFileAttachmentAsync('https://contoso.com/rtm/icon.png', 'icon.png', options, callback);
```

Vous pouvez vérifier la réussite ou l’échec d’un appel de méthode asynchrone dans la méthode de rappel à l’aide des propriétés `status` et `error` de l’objet `AsyncResult`. Si l’ajout de pièce jointe aboutit, vous pouvez utiliser la propriété `AsyncResult.value` pour obtenir l’ID de la pièce jointe. Il s’agit d’un nombre entier que vous pouvez ensuite utiliser pour supprimer la pièce jointe.

> [!NOTE]
> L’ID de pièce jointe n’est valide que dans la même session et il n’est pas garanti qu’il soit map enfant à la même pièce jointe entre les sessions. Les exemples de fin d’une session sont les suivants : lorsque l’utilisateur ferme le module, ou si l’utilisateur commence à composer dans un formulaire inline, puis ouvre le formulaire en ligne pour continuer dans une fenêtre distincte.

### <a name="attach-a-file"></a>Joindre un fichier

Vous pouvez joindre un fichier à un message ou un rendez-vous dans un formulaire de composition à l’aide de la méthode et en spécifiant `addFileAttachmentAsync` l’URI du fichier. Vous pouvez également utiliser la `addFileAttachmentFromBase64Async` méthode, mais spécifier la chaîne base64 comme entrée. Si le fichier est protégé, vous pouvez inclure une identité appropriée ou un jeton d’authentification comme paramètre de chaîne de requête d’URI. Exchange effectuera un appel à l’URI pour obtenir la pièce jointe, et le service web qui protège le fichier devra utiliser le jeton comme moyen d’authentification.

L’exemple JavaScript suivant est un complément de composition qui joint un fichier, picture.png, au message ou au rendez-vous en cours de composition à partir d’un serveur web. La méthode de rappel prend `asyncResult` comme paramètre, vérifie le statut du résultat et obtient l’ID de pièce jointe si la méthode a abouti.

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
                if (asyncResult.status === Office.AsyncResultStatus.Failed){
                    write(asyncResult.error.message);
                } else {
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

### <a name="attach-an-outlook-item"></a>Joindre un Outlook’élément

Vous pouvez joindre un élément Outlook (par exemple, un e-mail, un calendrier ou un élément de contact) à un message ou un rendez-vous dans un formulaire de composition en spécifiant l’ID Exchange Web Services (EWS) de l’élément et en utilisant la `addItemAttachmentAsync` méthode. Vous pouvez obtenir l’ID EWS d’un élément de messagerie, de calendrier, de contact ou de tâche dans la boîte aux lettres de l’utilisateur à l’aide de la méthode [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) et en accédant à l’opération EWS [FindItem](/exchange/client-developer/web-service-reference/finditem-operation). La propriété [item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) fournit également l’ID EWS d’un élément existant dans un formulaire de lecture.

La fonction JavaScript suivante, étend le premier exemple ci-dessus et ajoute un élément en tant que pièce jointe au message électronique ou au rendez-vous en `addItemAttachment` cours de composition. La fonction prend comme argument l’ID EWS de l’élément qui doit être joint. Si l’attachement réussit, il obtient l’ID de pièce jointe pour un traitement ultérieur, y compris la suppression de cette pièce jointe dans la même session.

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
            if (asyncResult.status === Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            } else {
                var attachmentID = asyncResult.value;
                write('ID of added attachment: ' + attachmentID);
            }
        });
}
```

> [!NOTE]
> Vous pouvez utiliser un module de composition pour joindre une instance d’un rendez-vous Outlook sur le web ou sur des appareils mobiles. Toutefois, dans une prise en charge Outlook client de bureau, la tentative d’attachement d’une instance entraînerait l’attachement de la série périodique (le rendez-vous parent).

## <a name="get-attachments"></a>Obtention de pièces jointes

Les API pour obtenir des pièces jointes en mode composition sont disponibles à partir de l’ensemble de conditions [requises 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).

- [getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)
- [getAttachmentContentAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

Vous pouvez utiliser la [méthode getAttachmentsAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) pour obtenir les pièces jointes du message ou du rendez-vous en cours de composition.

Pour obtenir le contenu d’une pièce jointe, vous pouvez utiliser la [méthode getAttachmentContentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Les formats pris en charge sont répertoriés dans [l’énumérer AttachmentContentFormat.](/javascript/api/outlook/office.mailboxenums.attachmentcontentformat)

Vous devez fournir une méthode de rappel pour vérifier l’état et toute erreur à l’aide de `AsyncResult` l’objet paramètre de sortie. Vous pouvez également transmettre des paramètres supplémentaires à la méthode de rappel à l’aide du paramètre `asyncContext` facultatif.

L’exemple JavaScript suivant obtient les pièces jointes et vous permet de configurer une gestion distincte pour chaque format de pièce jointe pris en charge.

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

## <a name="remove-an-attachment"></a>Supprimer une pièce jointe

Vous pouvez supprimer un fichier ou une pièce jointe d’un élément de message ou de rendez-vous dans un formulaire de composition en spécifiant l’ID de pièce jointe correspondant lors de l’utilisation de la méthode [removeAttachmentAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods)

> [!IMPORTANT]
> Si vous utilisez l’ensemble de conditions requises 1.7 ou une précédente, vous devez supprimer uniquement les pièces jointes ajoutées par le même module dans la même session.

Similaire à la méthode , et aux `addFileAttachmentAsync` `addItemAttachmentAsync` `getAttachmentsAsync` méthodes, `removeAttachmentAsync` est une méthode asynchrone. Vous devez fournir une méthode de rappel pour vérifier l’état et toute erreur à l’aide de `AsyncResult` l’objet paramètre de sortie. Vous pouvez également transmettre des paramètres supplémentaires à la méthode de rappel à l’aide du paramètre `asyncContext` facultatif.

La fonction JavaScript suivante, continue d’étendre les exemples ci-dessus et supprime la pièce jointe spécifiée de l’e-mail ou du rendez-vous en `removeAttachment` cours de composition. La fonction prend comme argument l’ID de la pièce jointe à supprimer. Vous pouvez obtenir l’ID d’une pièce jointe après un appel réussi, ou de méthode, et l’utiliser dans `addFileAttachmentAsync` un appel de méthode `addFileAttachmentFromBase64Async` `addItemAttachmentAsync` `removeAttachmentAsync` ultérieur. Vous pouvez également appeler (introduit dans l’ensemble de conditions requises 1.8) pour obtenir les pièces jointes et leurs ID pour cette `getAttachmentsAsync` session de module.

```js
// Removes the specified attachment from the composed item.
function removeAttachment(attachmentId) {
    // When the attachment is removed, the callback method is invoked.
    // Here, the callback method uses an asyncResult parameter and
    // gets the ID of the removed attachment if the removal succeeds.
    // You can optionally pass any object you wish to access in the
    // callback method as an argument to the asyncContext parameter.
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
