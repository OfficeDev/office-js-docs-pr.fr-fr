---
title: Insérer des données dans le corps dans un complément Outlook
description: Découvrez comment insérer des données dans le corps d’un message ou d’un rendez-vous dans un complément Outlook.
ms.date: 04/15/2019
localization_priority: Normal
ms.openlocfilehash: e092a67f8794c2821167ced84bede70a601c77e1
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324953"
---
# <a name="insert-data-in-the-body-when-composing-an-appointment-or-message-in-outlook"></a><span data-ttu-id="525d9-103">Insérer des données dans le corps lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="525d9-103">Insert data in the body when composing an appointment or message in Outlook</span></span>

<span data-ttu-id="525d9-p101">Vous pouvez utiliser les méthodes asynchrones ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) et [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) pour obtenir le type de corps et insérer des données dans le corps d’un élément de rendez-vous ou de message en cours de composition par l’utilisateur. Ces méthodes asynchrones sont disponibles uniquement pour les compléments de composition. Pour utiliser ces méthodes, assurez-vous que vous avez correctement défini le manifeste du complément afin qu’Outlook active le complément dans les formulaires de composition, comme décrit dans la rubrique [Créer des compléments Outlook pour les formulaires de composition](compose-scenario.md).</span><span class="sxs-lookup"><span data-stu-id="525d9-p101">You can use the asynchronous methods ([Body.getAsync](/javascript/api/outlook/office.Body#getasync-coerciontype--options--callback-), [Body.getTypeAsync](/javascript/api/outlook/office.Body#gettypeasync-options--callback-), [Body.prependAsync](/javascript/api/outlook/office.Body#prependasync-data--options--callback-), [Body.setAsync](/javascript/api/outlook/office.Body#setasync-data--options--callback-) and [Body.setSelectedDataAsync](/javascript/api/outlook/office.Body#setselecteddataasync-data--options--callback-)) to get the body type and insert data in the body of an appointment or message item that the user is composing. These asynchronous methods are available to only compose add-ins. To use these methods, make sure you have set up the add-in manifest appropriately so that Outlook activates your add-in in compose forms, as described in [Create Outlook add-ins for compose forms](compose-scenario.md).</span></span>

<span data-ttu-id="525d9-p102">Dans Outlook, un utilisateur peut créer un message au format texte, HTML ou RTF, ainsi qu’un rendez-vous au format HTML. Avant l’insertion, il est recommandé de d’abord vérifier le format de l’élément pris en charge en appelant la méthode **getTypeAsync**, car il est possible que vous ayez à suivre des étapes supplémentaires. La valeur que **getTypeAsync** renvoie dépend du format d’origine de l’élément, ainsi que de la prise en charge du système d’exploitation du dispositif et de l’hôte pour la modification au format HTML (1). Définissez ensuite le paramètre _coercionType_ des méthodes **prependAsync** ou **setSelectedDataAsync** en conséquence (2) pour insérer les données, tel qu’illustré dans le tableau ci-dessous. Si vous n’indiquez aucun argument, **prependAsync** et **setSelectedDataAsync** supposent que les données à insérer sont au format texte.</span><span class="sxs-lookup"><span data-stu-id="525d9-p102">In Outlook, a user can create a message in text, HTML, or Rich Text Format (RTF), and can create an appointment in HTML format. Before inserting, you should always first verify the supported item format by calling **getTypeAsync**, as you may need to take additional steps. The value that **getTypeAsync** returns depends on the original item format, as well as the support of the device operating system and host to editing in HTML format (1). Then set the  _coercionType_ parameter of **prependAsync** or **setSelectedDataAsync** accordingly (2) to insert the data, as shown in the following table. If you don't specify an argument, **prependAsync** and **setSelectedDataAsync** assume the data to insert is in text format.</span></span>

<br/>

|<span data-ttu-id="525d9-111">**Données à insérer**</span><span class="sxs-lookup"><span data-stu-id="525d9-111">**Data to insert**</span></span>|<span data-ttu-id="525d9-112">**Format de l’élément retourné par getTypeAsync**</span><span class="sxs-lookup"><span data-stu-id="525d9-112">**Item format returned by getTypeAsync**</span></span>|<span data-ttu-id="525d9-113">**Utiliser ce paramètre coercionType**</span><span class="sxs-lookup"><span data-stu-id="525d9-113">**Use this coercionType**</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="525d9-114">Texte</span><span class="sxs-lookup"><span data-stu-id="525d9-114">Text</span></span>|<span data-ttu-id="525d9-115">Texte (1)</span><span class="sxs-lookup"><span data-stu-id="525d9-115">Text (1)</span></span>|<span data-ttu-id="525d9-116">Texte</span><span class="sxs-lookup"><span data-stu-id="525d9-116">Text</span></span>|
|<span data-ttu-id="525d9-117">HTML</span><span class="sxs-lookup"><span data-stu-id="525d9-117">HTML</span></span>|<span data-ttu-id="525d9-118">Texte (1)</span><span class="sxs-lookup"><span data-stu-id="525d9-118">Text (1)</span></span>|<span data-ttu-id="525d9-119">Texte (2)</span><span class="sxs-lookup"><span data-stu-id="525d9-119">Text (2)</span></span>|
|<span data-ttu-id="525d9-120">Texte</span><span class="sxs-lookup"><span data-stu-id="525d9-120">Text</span></span>|<span data-ttu-id="525d9-121">HTML</span><span class="sxs-lookup"><span data-stu-id="525d9-121">HTML</span></span>|<span data-ttu-id="525d9-122">Texte/HTML</span><span class="sxs-lookup"><span data-stu-id="525d9-122">Text/HTML</span></span>|
|<span data-ttu-id="525d9-123">HTML</span><span class="sxs-lookup"><span data-stu-id="525d9-123">HTML</span></span>|<span data-ttu-id="525d9-124">HTML</span><span class="sxs-lookup"><span data-stu-id="525d9-124">HTML</span></span> |<span data-ttu-id="525d9-125">HTML</span><span class="sxs-lookup"><span data-stu-id="525d9-125">HTML</span></span>|

1.  <span data-ttu-id="525d9-126">Sur les tablettes et les smartphones, la méthode **getTypeAsync** renvoie **Office.MailboxEnums.BodyType.Text** si le système d’exploitation ou l’hôte ne prend pas en charge la modification d’un élément qui a été créé à l’origine au format HTML.</span><span class="sxs-lookup"><span data-stu-id="525d9-126">On tablets and smartphones, **getTypeAsync** returns **Office.MailboxEnums.BodyType.Text** if the operating system or host does not support editing an item, which was originally created in HTML, in HTML format.</span></span>

2.  <span data-ttu-id="525d9-p103">Si les données à insérer sont au format HTML et que la méthode **getTypeAsync** renvoie un type de texte pour cet élément, réorganisez vos données au format texte et insérez-les avec **Office.MailboxEnums.BodyType.Text** en tant que _coercionType_. Si vous insérez simplement les données HTML avec un type de forçage de type texte, l’hôte va afficher les balises HTML comme du texte. Si vous essayez d’insérer les données HTML avec **Office.MailboxEnums.BodyType.Html** en tant que _coercionType_, vous obtenez une erreur.</span><span class="sxs-lookup"><span data-stu-id="525d9-p103">If your data to insert is HTML and **getTypeAsync** returns a text type for that item, reorganize your data as text and insert it with **Office.MailboxEnums.BodyType.Text** as _coercionType_. If you simply insert the HTML data with a text coercion type, the host would display the HTML tags as text. If you attempt to insert the HTML data with **Office.MailboxEnums.BodyType.Html** as _coercionType_, you will get an error.</span></span>

<span data-ttu-id="525d9-p104">En plus de _coercionType_, comme pour la plupart des méthodes asynchrones dans l’API JavaScript pour Office, **getTypeAsync**, **prependAsync** et **setSelectedDataAsync** prennent d’autres paramètres d’entrée facultatifs. Pour plus d’informations sur la spécification de ces paramètres d’entrée facultatifs, consultez la rubrique [passing Optional Parameters to Asynchronous Methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous Programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="525d9-p104">In addition to  _coercionType_, as with most asynchronous methods in the Office JavaScript API, **getTypeAsync**, **prependAsync** and **setSelectedDataAsync** take other optional input parameters. For more information about specifying these optional input parameters, see [passing optional parameters to asynchronous methods](../develop/asynchronous-programming-in-office-add-ins.md#passing-optional-parameters-inline) in [Asynchronous programming in Office Add-ins](../develop/asynchronous-programming-in-office-add-ins.md).</span></span>


## <a name="insert-data-at-the-current-cursor-position"></a><span data-ttu-id="525d9-132">Insertion de données à l’emplacement du curseur</span><span class="sxs-lookup"><span data-stu-id="525d9-132">Insert data at the current cursor position</span></span>


<span data-ttu-id="525d9-133">Cette section présente un exemple de code qui utilise la méthode **getTypeAsync** pour vérifier le type de corps de l’élément dont la composition est en cours, puis la méthode **setSelectedDataAsync** pour insérer des données à l’emplacement du curseur.</span><span class="sxs-lookup"><span data-stu-id="525d9-133">This section shows a code sample that uses **getTypeAsync** to verify the body type of the item that is being composed, and then uses **setSelectedDataAsync** to insert data in the current cursor location.</span></span>

<span data-ttu-id="525d9-p105">Vous pouvez transmettre une méthode de rappel et ses paramètres d’entrée facultatifs à la méthode **getTypeAsync**, et obtenir le statut et les résultats dans le paramètre de sortie _asyncResult_. Si la méthode aboutit, vous pouvez obtenir le type de corps de l’élément dans la propriété [AsyncResult.value](/javascript/api/office/office.asyncresult#value), à savoir « text » ou « html ».</span><span class="sxs-lookup"><span data-stu-id="525d9-p105">You can pass a callback method and optional input parameters to **getTypeAsync**, and get any status and results in the  _asyncResult_ output parameter. If the method succeeds, you can get the type of the item body in the [AsyncResult.value](/javascript/api/office/office.asyncresult#value) property, which is either "text" or "html".</span></span>

<span data-ttu-id="525d9-p106">Vous devez transmettre une chaîne de données comme paramètre d’entrée à la méthode **setSelectedDataAsync**. Selon le type de corps de l’élément, vous pouvez spécifier cette chaîne de données au format texte ou HTML. Comme mentionné ci-dessus, vous pouvez éventuellement spécifier le type de données à insérer dans le paramètre _coercionType_. En outre, vous pouvez fournir une méthode de rappel et ses paramètres comme paramètres d’entrée facultatifs.</span><span class="sxs-lookup"><span data-stu-id="525d9-p106">You must pass a data string as an input parameter to **setSelectedDataAsync**. Depending on the type of the item body, you can specify this data string in text or HTML format accordingly. As mentioned above, you can optionally specify the type of the data to be inserted in the  _coercionType_ parameter. In addition, you can provide a callback method and any of its parameters as optional input parameters.</span></span>

<span data-ttu-id="525d9-p107">Si l’utilisateur n’a pas placé le curseur dans le corps de l’élément, **setSelectedDataAsync** insère les données au début du corps de l’élément. Si l’utilisateur a sélectionné du texte dans le corps de l’élément, **setSelectedDataAsync** remplace le texte sélectionné par les données spécifiées. Notez que la méthode **setSelectedDataAsync** peut échouer si l’utilisateur change l’emplacement du curseur lors de la composition de l’élément. Vous pouvez insérer simultanément jusqu’à 1 000 000 caractères.</span><span class="sxs-lookup"><span data-stu-id="525d9-p107">If the user hasn't placed the cursor in the item body, **setSelectedDataAsync** inserts the data at the top of the body. If the user has selected text in the item body, **setSelectedDataAsync** replaces the selected text by the data you specify. Note that **setSelectedDataAsync** can fail if the user is simultaneously changing the cursor position while composing the item. The maximum number of characters you can insert at one time is 1,000,000 characters.</span></span>

<span data-ttu-id="525d9-144">Cet exemple de code suppose l’existence d’une règle dans le manifeste du complément qui active le complément dans un formulaire de composition pour un rendez-vous ou un message, comme indiqué ci-dessous.</span><span class="sxs-lookup"><span data-stu-id="525d9-144">This code sample assumes a rule in the add-in manifest that activates the add-in in a compose form for an appointment or message, as shown below.</span></span>




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


## <a name="insert-data-at-the-beginning-of-the-item-body"></a><span data-ttu-id="525d9-145">Insertion de données au début du corps de l’élément</span><span class="sxs-lookup"><span data-stu-id="525d9-145">Insert data at the beginning of the item body</span></span>


<span data-ttu-id="525d9-p108">Vous pouvez également utiliser la méthode **prependAsync** pour insérer des données au début du corps de l’élément et ne pas tenir compte de l’emplacement du curseur. Mis à part le point d’insertion, les méthodes **prependAsync** et **setSelectedDataAsync** se comportent de façon similaire :</span><span class="sxs-lookup"><span data-stu-id="525d9-p108">Alternatively, you can use **prependAsync** to insert data at the beginning of the item body and disregard the current cursor location. Other than the point of insertion, **prependAsync** and **setSelectedDataAsync** behave in similar ways:</span></span>


- <span data-ttu-id="525d9-148">Si vous ajoutez des données HTML dans le corps d’un message, vous devez d’abord vérifier le type de corps du message pour éviter d’ajouter des données HTML dans un message au format texte.</span><span class="sxs-lookup"><span data-stu-id="525d9-148">If you are prepending HTML data in a message body, you should first check for the type of the message body to avoid prepending HTML data to a message in text format.</span></span>
    
- <span data-ttu-id="525d9-149">Fournissez les éléments suivants comme paramètres d’entrée dans la méthode **prependAsync** : une chaîne de données au format texte ou HTML et éventuellement le format des données à insérer, une méthode de rappel et ses paramètres.</span><span class="sxs-lookup"><span data-stu-id="525d9-149">Provide the following as input parameters to **prependAsync**: a data string in either text or HTML format, and optionally the format of the data to be inserted, a callback method and any of its parameters.</span></span>
    
- <span data-ttu-id="525d9-150">Vous pouvez ajouter simultanément jusqu’à 1 000 000 caractères.</span><span class="sxs-lookup"><span data-stu-id="525d9-150">The maximum number of characters you can prepend at one time is 1,000,000 characters.</span></span>
    
<span data-ttu-id="525d9-p109">Le code JavaScript suivant fait partie d’un exemple de complément activé dans les formulaires de composition de rendez-vous et de messages. L’exemple appelle la méthode **getTypeAsync** pour vérifier le type de corps de l’élément. Il insère ensuite les données HTML au début du corps de l’élément si ce dernier est un rendez-vous ou un message HTML ; dans le cas contraire, il insère les données au format texte.</span><span class="sxs-lookup"><span data-stu-id="525d9-p109">The following JavaScript code is part of a sample add-in that is activated in compose forms of appointments and messages. The sample calls **getTypeAsync** to verify the type of the item body, inserts HTML data to the top of the item body if the item is an appointment or HTML message, otherwise inserts the data in text format.</span></span>




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


## <a name="see-also"></a><span data-ttu-id="525d9-153">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="525d9-153">See also</span></span>

- [<span data-ttu-id="525d9-154">Obtenir et définir des données d’élément dans un formulaire de composition dans Outlook</span><span class="sxs-lookup"><span data-stu-id="525d9-154">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)    
- [<span data-ttu-id="525d9-155">Obtenir et définir des données d’élément Outlook dans des formulaires de lecture ou de composition</span><span class="sxs-lookup"><span data-stu-id="525d9-155">Get and set Outlook item data in read or compose forms</span></span>](item-data.md)    
- [<span data-ttu-id="525d9-156">Créer des compléments Outlook pour les formulaires de composition</span><span class="sxs-lookup"><span data-stu-id="525d9-156">Create Outlook add-ins for compose forms</span></span>](compose-scenario.md)    
- [<span data-ttu-id="525d9-157">Programmation asynchrone dans les compléments Office</span><span class="sxs-lookup"><span data-stu-id="525d9-157">Asynchronous programming in Office Add-ins</span></span>](../develop/asynchronous-programming-in-office-add-ins.md)    
- [<span data-ttu-id="525d9-158">Obtenir, définir ou ajouter des destinataires lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="525d9-158">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)  
- [<span data-ttu-id="525d9-159">Obtenir ou définir l’objet lors de la composition d’un rendez-vous ou d’un message dans Outlook</span><span class="sxs-lookup"><span data-stu-id="525d9-159">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)  
- [<span data-ttu-id="525d9-160">Obtenir ou définir l’emplacement lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="525d9-160">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md) 
- [<span data-ttu-id="525d9-161">Obtenir ou définir l’heure lors de la composition d’un rendez-vous dans Outlook</span><span class="sxs-lookup"><span data-stu-id="525d9-161">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)
    
