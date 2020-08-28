---
title: Insérer des données dans le corps dans un complément Outlook
description: Découvrez comment insérer des données dans le corps d’un message ou d’un rendez-vous dans un complément Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: 0e875619520ee309dec97b2db60ed49c29b2a463
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293869"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a>Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook

Vous pouvez utiliser les méthodes asynchrones ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) et [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) pour obtenir le type de corps et insérer des données dans le corps d’un élément de rendez-vous ou de message en cours de composition par l’utilisateur. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, assurez-vous que vous avez correctement défini le manifeste du complément afin qu’Outlook active le complément dans les formulaires de composition, comme décrit dans la rubrique [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md).

Dans Outlook, un utilisateur peut créer un message au format texte, HTML ou RTF (Rich Text Format) et il peut créer un rendez-vous au format HTML. Avant l’insertion, vous devez toujours vérifier d’abord le format d’élément pris en charge en appelant **getTypeAsync**, car vous devrez peut-être effectuer des étapes supplémentaires. La valeur renvoyée par **getTypeAsync** dépend du format d’élément d’origine, ainsi que de la prise en charge du système d’exploitation de l’appareil et de l’application à modifier au format html (1). Définissez ensuite le paramètre  _coercionType_ de **prependAsync** ou **setSelectedDataAsync** en conséquence (2) pour insérer les données, comme indiqué dans le tableau suivant. Si vous ne spécifiez pas d’argument, **prependAsync** et **setSelectedDataAsync** supposent que les données à insérer sont au format texte.

<br/>

|**Données à insérer**|**Format de l’élément retourné par getTypeAsync**|**Utiliser ce paramètre coercionType**|
|:-----|:-----|:-----|
|Texte|Texte (1)|Texte|
|HTML|Texte (1)|Texte (2)|
|Texte|HTML|Texte/HTML|
|HTML|HTML |HTML|

1.  Sur les tablettes et les smartphones, **getTypeAsync** renvoie **Office. MailboxEnums. BodyType. Text** si le système d’exploitation ou l’application ne prend pas en charge la modification d’un élément qui a été initialement créé au format HTML au format html.

2.  Si vos données à insérer sont au format HTML et que **getTypeAsync** renvoie un type de texte pour cet élément, réorganisez vos données en tant que texte et insérez-le avec **Office. MailboxEnums. BodyType. Text** en tant que _coercionType_. Si vous insérez simplement les données HTML avec un type de forçage de texte, l’application affiche les balises HTML sous forme de texte. Si vous tentez d’insérer les données HTML avec **Office.MailboxEnums.BodyType.Html** en tant que _coercionType_, vous obtiendrez une erreur.

En plus de  _coercionType_, comme pour la plupart des méthodes asynchrones dans l’API JavaScript pour Office, **getTypeAsync**, **prependAsync** et **setSelectedDataAsync** prennent d’autres paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, consultez la rubrique [passing Optional Parameters to Asynchronous Methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous Programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).


## <a name="insert-data-at-the-current-cursor-position"></a>Insertion de données à l’emplacement du curseur


Cette section présente un exemple de code qui utilise la méthode **getTypeAsync** pour vérifier le type de corps de l’élément dont la composition est en cours, puis la méthode **setSelectedDataAsync** pour insérer des données à l’emplacement du curseur.

Vous pouvez transmettre une méthode de rappel et ses paramètres d’entrée facultatifs à la méthode **getTypeAsync**, et obtenir le statut et les résultats dans le paramètre de sortie _asyncResult_. Si la méthode aboutit, vous pouvez obtenir le type de corps de l’élément dans la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value), à savoir « text » ou « html ».

Vous devez transmettre une chaîne de données comme paramètre d’entrée à la méthode **setSelectedDataAsync**. Selon le type de corps de l’élément, vous pouvez spécifier cette chaîne de données au format texte ou HTML. Comme mentionné ci-dessus, vous pouvez éventuellement spécifier le type de données à insérer dans le paramètre _coercionType_. En outre, vous pouvez fournir une méthode de rappel et ses paramètres comme paramètres d’entrée facultatifs.

Si l’utilisateur n’a pas placé le curseur dans le corps de l’élément, **setSelectedDataAsync** insère les données au début du corps de l’élément. Si l’utilisateur a sélectionné du texte dans le corps de l’élément, **setSelectedDataAsync** remplace le texte sélectionné par les données spécifiées. Notez que la méthode **setSelectedDataAsync** peut échouer si l’utilisateur change l’emplacement du curseur lors de la composition de l’élément. Vous pouvez insérer simultanément jusqu’à 1 000 000 caractères.

Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous.




```XML
<Rule xsi:type="RuleCollection" Mode="Or">
  <Rule xsi:type="ItemIs" ItemType="Appointment" FormType="Edit"/>
  <Rule xsi:type="ItemIs" ItemType="Message" FormType="Edit"/>
</Rule>

```




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Set data in the body of the composed item.
        setItemBody();
    });
}


// Get the body type of the composed item, and set data in 
// in the appropriate data type in the item body.
function setItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Set data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of setSelectedDataAsync.
                    item.body.setSelectedDataAsync(
                        '<b> Kindly note we now open 7 days a week.</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.setSelectedDataAsync(
                        ' Kindly note we now open 7 days a week.',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully set data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="insert-data-at-the-beginning-of-the-item-body"></a>Insertion de données au début du corps de l’élément


Vous pouvez également utiliser la méthode **prependAsync** pour insérer des données au début du corps de l’élément et ne pas tenir compte de l’emplacement du curseur. Mis à part le point d’insertion, les méthodes **prependAsync** et **setSelectedDataAsync** se comportent de façon similaire :


- Si vous ajoutez des données HTML dans le corps d’un message, vous devez d’abord vérifier le type de corps du message pour éviter d’ajouter des données HTML dans un message au format texte.
    
- Fournissez les éléments suivants comme paramètres d’entrée dans la méthode **prependAsync** : une chaîne de données au format texte ou HTML et éventuellement le format des données à insérer, une méthode de rappel et ses paramètres.
    
- Vous pouvez ajouter simultanément jusqu’à 1 000 000 caractères.
    
Le code JavaScript suivant fait partie d’un exemple de complément activé dans les formulaires de composition de rendez-vous et de messages. L’exemple appelle la méthode **getTypeAsync** pour vérifier le type de corps de l’élément. Il insère ensuite les données HTML au début du corps de l’élément si ce dernier est un rendez-vous ou un message HTML ; dans le cas contraire, il insère les données au format texte.




```js
var item;

Office.initialize = function () {
    item = Office.context.mailbox.item;
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
        // Insert data in the top of the body of the composed 
        // item.
        prependItemBody();
    });
}

// Get the body type of the composed item, and prepend data  
// in the appropriate data type in the item body.
function prependItemBody() {
    item.body.getTypeAsync(
        function (result) {
            if (result.status == Office.AsyncResultStatus.Failed){
                write(asyncResult.error.message);
            }
            else {
                // Successfully got the type of item body.
                // Prepend data of the appropriate type in body.
                if (result.value == Office.MailboxEnums.BodyType.Html) {
                    // Body is of HTML type.
                    // Specify HTML in the coercionType parameter
                    // of prependAsync.
                    item.body.prependAsync(
                        '<b>Greetings!</b>',
                        { coercionType: Office.CoercionType.Html, 
                        asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                        });
                }
                else {
                    // Body is of text type. 
                    item.body.prependAsync(
                        'Greetings!',
                        { coercionType: Office.CoercionType.Text, 
                            asyncContext: { var3: 1, var4: 2 } },
                        function (asyncResult) {
                            if (asyncResult.status == 
                                Office.AsyncResultStatus.Failed){
                                write(asyncResult.error.message);
                            }
                            else {
                                // Successfully prepended data in item body.
                                // Do whatever appropriate for your scenario,
                                // using the arguments var3 and var4 as applicable.
                            }
                         });
                }
            }
        });

}

// Writes to a div with id='message' on the page.
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
- [Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-location-of-an-appointment.md) 
- [Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook](get-or-set-the-time-of-an-appointment.md)
    
