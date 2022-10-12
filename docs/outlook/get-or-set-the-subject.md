---
title: Obtenir ou définir l’objet dans un complément Outlook
description: Découvrez comment obtenir ou définir l’objet d’un message ou d’un rendez-vous dans un complément Outlook.
ms.date: 10/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 79e38a310bf62eae55ef020c2f6c978ace824255
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541128"
---
# <a name="get-or-set-the-subject-when-composing-an-appointment-or-message-in-outlook"></a>Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook

L’API JavaScript Office fournit des méthodes asynchrones ([subject.getAsync](/javascript/api/outlook/office.subject#outlook-office-subject-getasync-member(1)) et [subject.setAsync](/javascript/api/outlook/office.subject#outlook-office-subject-setasync-member(1))) pour obtenir et définir l’objet d’un rendez-vous ou d’un message que l’utilisateur compose. Ces méthodes asynchrones sont disponibles uniquement pour composer des compléments. Pour utiliser ces méthodes, assurez-vous que vous avez correctement configuré le manifeste XML du complément pour qu’Outlook [active le complément dans les formulaires de composition](compose-scenario.md). Les règles d’activation ne sont pas prises en charge dans les compléments qui utilisent un [manifeste Teams pour les compléments Office (préversion).](../develop/json-manifest-overview.md)

The **subject** property is available for read access in both compose and read forms of appointments and messages. In a read form, you can access the property directly from the parent object, as in:

```js
item.subject
```

Cependant, dans un formulaire de composition, comme l’utilisateur et votre complément peuvent insérer ou modifier l’objet en même temps, vous devez utiliser la méthode asynchrone **getAsync** pour obtenir l’objet, comme indiqué ci-dessous :

```js
item.subject.getAsync
```

La propriété **subject** est disponible pour l’accès en écriture uniquement dans les formulaires de composition, pas dans les formulaires de lecture.

Comme avec la plupart des méthodes asynchrones dans l’API JavaScript Office, **getAsync** et **setAsync prennent des paramètres d’entrée** facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, voir la section « Passage de paramètres facultatifs à des méthodes asynchrones » dans la rubrique [Programmation asynchrone dans des compléments Office](../develop/asynchronous-programming-in-office-add-ins.md).

## <a name="get-the-subject"></a>Obtention de l’objet

Cette section présente un exemple de code qui obtient l’objet du rendez-vous ou du message que l’utilisateur compose, et affiche l’objet. Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous. Les règles d’activation ne sont pas prises en charge dans les compléments qui utilisent un [manifeste Teams pour les compléments Office (préversion).](../develop/json-manifest-overview.md)

```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>
```

Pour utiliser **item.subject.getAsync**, fournissez une fonction de rappel qui vérifie l’état et le résultat de l’appel asynchrone. Vous pouvez fournir tous les arguments nécessaires à la fonction de rappel via le paramètre facultatif  _asyncContext_ . Vous pouvez obtenir l’état, les résultats et toute erreur à l’aide du paramètre de sortie  _asyncResult_ du rappel. Si l’appel asynchrone aboutit, vous pouvez obtenir l’objet sous forme de chaîne de texte brut à l’aide de la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#office-office-asyncresult-value-member).

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
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

Cette section présente un exemple de code qui définit l’objet du rendez-vous ou du message que l’utilisateur compose. Comme dans l’exemple précédent, cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message. Les règles d’activation ne sont pas prises en charge dans les compléments qui utilisent un [manifeste Teams pour les compléments Office (préversion).](../develop/json-manifest-overview.md)

Pour utiliser **item.subject.setAsync**, spécifiez une chaîne de 255 caractères maximum dans le paramètre de données. Si vous le souhaitez, vous pouvez fournir une fonction de rappel et tous les arguments de la fonction de rappel dans le paramètre  _asyncContext_ . Vous devez vérifier l’état, le résultat et tous les messages d’erreur dans le paramètre de sortie  _asyncResult_ du rappel. Si l’appel asynchrone aboutit, **setAsync** insère la chaîne d’objet spécifiée sous forme de texte brut, en écrasant tous les objets existants pour cet élément.

```js
let item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready method.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set the subject of the item being composed.
        setSubject();
    });
}

// Set the subject of the item that the user is composing.
function setSubject() {
    const today = new Date();
    let subject;

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
