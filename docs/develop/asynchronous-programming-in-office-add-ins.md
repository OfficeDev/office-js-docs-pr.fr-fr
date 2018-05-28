---
title: Programmation asynchrone dans des compl?ments Office
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: d251ebfd03227569b9a24bcd7f17baada6099938
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 05/23/2018
---
# <a name="asynchronous-programming-in-office-add-ins"></a><span data-ttu-id="a9446-102">Programmation asynchrone dans des compl?ments Office</span><span class="sxs-lookup"><span data-stu-id="a9446-102">Asynchronous programming in Office Add-ins</span></span>

<span data-ttu-id="a9446-p101">Pourquoi l?API de Compl?ments Office a-t-elle recours ? la programmation asynchrone ?JavaScript ?tant un langage monothread, si le script appelle un processus synchrone de longue dur?e, toute ex?cution de script ult?rieure sera bloqu?e tant que ce processus ne sera pas termin?. Comme certaines op?rations, notamment celles agissant sur les clients web Office (mais aussi sur les clients riches), peuvent bloquer l?ex?cution si elles sont ex?cut?es de fa?on synchrone, la plupart des m?thodes dans l?interface API JavaScript pour Office sont con?ues pour ?tre ex?cut?es de fa?on asynchrone. Cela permet de garantir que les Compl?ments Office sont r?actifs et tr?s performants. Vous devez donc fr?quemment ?crire des fonctions de rappel lorsque vous utilisez ces m?thodes asynchrones.</span><span class="sxs-lookup"><span data-stu-id="a9446-p101">Why does the Office Add-ins API use asynchronous programming? Because JavaScript is a single-threaded language, if script invokes a long-running synchronous process, all subsequent script execution will be blocked until that process completes. Because certain operations against Office web clients (but rich clients as well) could block execution if they are run synchronously, most of the methods in the JavaScript API for Office are designed to execute asynchronously. This makes sure that Office Add-ins are responsive and highly performing. It also frequently requires you to write callback functions when working with these asynchronous methods.</span></span>

<span data-ttu-id="a9446-p102">Le nom de toutes les m?thodes asynchrones de l?API se terminent par ? Async ?, comme pour les m?thodes [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) ou [Item.loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item). Lorsqu?une m?thode ? Async ? est appel?e, elle est ex?cut?e imm?diatement et toute ex?cution de script ult?rieure peut se poursuivre normalement. La fonction de rappel facultative que vous transmettez ? une m?thode ? Async ? s?ex?cute d?s que l?op?ration demand?e ou les donn?es sont pr?tes. L?op?ration est g?n?ralement rapide, mais le retour pourrait pr?senter un l?ger retard.</span><span class="sxs-lookup"><span data-stu-id="a9446-p102">The names of all asynchronous methods in the API end with "Async", such as the  [Document.getSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.getselecteddataasync), [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), or [Item.loadCustomPropertiesAsync](https://dev.office.com/reference/add-ins/outlook/Office.context.mailbox.item) methods. When an "Async" method is called, it executes immediately and any subsequent script execution can continue. The optional callback function you pass to an "Async" method executes as soon as the data or requested operation is ready. This generally occurs promptly, but there can be a slight delay before it returns.</span></span>

<span data-ttu-id="a9446-p103">Le diagramme suivant pr?sente le flux d?ex?cution d?un appel ? une m?thode ? Async ? qui lit les donn?es s?lectionn?es par l?utilisateur dans un document ouvert dans l?instance Word Online ou Excel Online sur le serveur. Au moment o? l?appel ? Async ? est effectu?, le thread d?ex?cution JavaScript est libre d?effectuer tout traitement c?t? client suppl?mentaire (m?me si aucun n?est affich? dans le diagramme). Lors du retour de la m?thode ? Async ?, l?appel reprend l?ex?cution sur le thread et le compl?ment peut acc?der aux donn?es, les exploiter et afficher le r?sultat. Le m?me motif d?ex?cution asynchrone est employ? en cas d?utilisation des applications h?tes de client riche Office, telles que Word 2013 ou Excel 2013.</span><span class="sxs-lookup"><span data-stu-id="a9446-p103">The following diagram shows the flow of execution for a call to an "Async" method that reads the data the user selected in a document open in the server-based Word Online or Excel Online. At the point when the "Async" call is made, the JavaScript execution thread is free to perform any additional client-side processing. (Although none are shown in the diagram.) When the "Async" method returns, the callback resumes execution on the thread, and the add-in can the access data, do something with it, and display the result. The same asynchronous execution pattern holds when working with the Office rich client host applications, such as Word 2013 or Excel 2013.</span></span>

<span data-ttu-id="a9446-116">*Figure 1. Flux d?ex?cution de programmation asynchrone*</span><span class="sxs-lookup"><span data-stu-id="a9446-116">*Figure 1. Asynchronous programing execution flow*</span></span>

![Flux d?ex?cution de thread de programmation asynchrone](../images/office15-app-async-prog-fig01.png)

<span data-ttu-id="a9446-p104">La prise en charge de cette conception asynchrone dans les clients riches et les clients web fait partie des objectifs de conception ? ?criture unique-ex?cution multiplateforme ? du mod?le de d?veloppement des Compl?ments Office. Par exemple, vous pouvez cr?er un compl?ment de contenu ou du volet de t?ches avec une seule base de code qui sera ex?cut?e sur Excel 2013 et Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a9446-p104">Support for this asynchronous design in both rich and web clients is part of the "write once-run cross-platform" design goals of the Office Add-ins development model. For example, you can create a content or task pane add-in with a single code base that will run in both Excel 2013 and Excel Online.</span></span>

## <a name="writing-the-callback-function-for-an-async-method"></a><span data-ttu-id="a9446-120">?criture de la fonction de rappel pour une m?thode ? Async ?</span><span class="sxs-lookup"><span data-stu-id="a9446-120">Writing the callback function for an "Async" method</span></span>


<span data-ttu-id="a9446-p105">La fonction de rappel que vous transmettez en tant qu?argument _callback_ ? une m?thode ? Async ? doit d?clarer un seul param?tre que le runtime de compl?ment va utiliser pour permettre l?acc?s ? un objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) lorsque la fonction de rappel sera ex?cut?e. Vous pouvez ?crire :</span><span class="sxs-lookup"><span data-stu-id="a9446-p105">The callback function you pass as the  _callback_ argument to an "Async" method must declare a single parameter that the add-in runtime will use to provide access to an [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object when the callback function executes. You can write:</span></span>


- <span data-ttu-id="a9446-123">une fonction anonyme devant ?tre ?crite et pass?e directement en ligne avec l?appel ? la m?thode ? Async ? en tant que param?tre  _callback_ de la m?thode ? Async ? ;</span><span class="sxs-lookup"><span data-stu-id="a9446-123">An anonymous function that must be written and passed directly in line with the call to the "Async" method as the  _callback_ parameter of the "Async" method.</span></span>
    
- <span data-ttu-id="a9446-124">une fonction nomm?e, en passant le nom de cette fonction en tant que param?tre  _callback_ de la m?thode ? Async ?.</span><span class="sxs-lookup"><span data-stu-id="a9446-124">A named function, passing the name of that function as the  _callback_ parameter of an "Async" method.</span></span>
    
<span data-ttu-id="a9446-p106">Une fonction anonyme est utile si vous envisagez de n?utiliser son code qu?une fois : comme elle n?a pas de nom, vous ne pouvez pas y faire r?f?rence dans une autre partie du code. Une fonction nomm?e est utile si vous voulez r?utiliser la fonction de rappel pour plusieurs m?thodes ? Async ?.</span><span class="sxs-lookup"><span data-stu-id="a9446-p106">An anonymous function is useful if you are only going to use its code once - because it has no name, you can't reference it in another part of your code. A named function is useful if you want to reuse the callback function for more than one "Async" method.</span></span>


### <a name="writing-an-anonymous-callback-function"></a><span data-ttu-id="a9446-127">?criture d?une fonction de rappel anonyme</span><span class="sxs-lookup"><span data-stu-id="a9446-127">Writing an anonymous callback function</span></span>

<span data-ttu-id="a9446-128">La fonction de rappel anonyme suivante d?clare un seul param?tre nomm? `result` qui r?cup?re les donn?es ? partir de la propri?t? [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) lorsque le rappel est renvoy?.</span><span class="sxs-lookup"><span data-stu-id="a9446-128">The following anonymous callback function declares a single parameter named  `result` that retrieves data from the [AsyncResult.value](https://dev.office.com/reference/add-ins/shared/asyncresult.status) property when the callback returns.</span></span>


```js
function (result) {
        write('Selected data: ' + result.value);
}
```

<span data-ttu-id="a9446-129">L?exemple suivant montre comment passer cette fonction de rappel anonyme dans le contexte d?un appel complet de m?thode ? Async ? ? la m?thode  **Document.getSelectedDataAsync**.</span><span class="sxs-lookup"><span data-stu-id="a9446-129">The following example shows how to pass this anonymous callback function in line in the context of a full "Async" method call to the  **Document.getSelectedDataAsync** method.</span></span>


- <span data-ttu-id="a9446-130">Le premier argument  _coercionType_,  `Office.CoercionType.Text`, sp?cifie le retour des donn?es s?lectionn?es en tant que cha?ne de texte.</span><span class="sxs-lookup"><span data-stu-id="a9446-130">The first  _coercionType_ argument, `Office.CoercionType.Text`, specifies to return the selected data as a string of text.</span></span>
    
- <span data-ttu-id="a9446-p107">Le deuxi?me argument  _callback_ est la fonction anonyme pass?e en ligne ? la m?thode. Lors de l?ex?cution de la fonction, elle utilise le param?tre _result_ pour acc?der ? la propri?t? **value** de l?objet **AsyncResult** afin d?afficher les donn?es s?lectionn?es par l?utilisateur dans le document.</span><span class="sxs-lookup"><span data-stu-id="a9446-p107">The second  _callback_ argument is the anonymous function passed in-line to the method. When the function executes, it uses the _result_ parameter to access the **value** property of the **AsyncResult** object to display the data selected by the user in the document.</span></span>
    



```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    function (result) {
        write('Selected data: ' + result.value);
    }
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="a9446-p108">Vous pouvez ?galement utiliser le param?tre de votre fonction de rappel pour acc?der aux autres propri?t?s de l?objet **AsyncResult**. Utilisez la propri?t? [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) pour d?terminer si l?appel a r?ussi ou ?chou?. En cas d??chec, vous pouvez utiliser la propri?t? [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) pour acc?der ? un objet [Error](https://dev.office.com/reference/add-ins/shared/error) et obtenir des informations sur l?erreur.</span><span class="sxs-lookup"><span data-stu-id="a9446-p108">You can also use the parameter of your callback function to access other properties of the  **AsyncResult** object. Use the [AsyncResult.status](https://dev.office.com/reference/add-ins/shared/asyncresult.error) property to determine if the call succeeded or failed. If your call fails you can use the [AsyncResult.error](https://dev.office.com/reference/add-ins/shared/asyncresult.context) property to access an [Error](https://dev.office.com/reference/add-ins/shared/error) object for error information.</span></span>

<span data-ttu-id="a9446-136">Pour plus d?informations sur l?utilisation de la m?thode  **getSelectedDataAsync**, voir [Lecture et ?criture de donn?es dans la s?lection active d?un document ou d?une feuille de calcul](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="a9446-136">For more information about using the  **getSelectedDataAsync** method, see [Read and write data to the active selection in a document or spreadsheet](read-and-write-data-to-the-active-selection-in-a-document-or-spreadsheet.md).</span></span> 


### <a name="writing-a-named-callback-function"></a><span data-ttu-id="a9446-137">?criture d?une fonction de rappel nomm?e</span><span class="sxs-lookup"><span data-stu-id="a9446-137">Writing a named callback function</span></span>

<span data-ttu-id="a9446-p109">Vous pouvez ?galement ?crire une fonction nomm?e et passer son nom au param?tre  _callback_ d?une m?thode ? Async ?. Par exemple, l?exemple pr?c?dent peut ?tre r??crit pour passer une fonction nomm?e `writeDataCallback` en tant que param?tre _callback_ comme suit.</span><span class="sxs-lookup"><span data-stu-id="a9446-p109">Alternatively, you can write a named function and pass its name to the  _callback_ parameter of an "Async" method. For example, the previous example can be rewritten to pass a function named `writeDataCallback` as the _callback_ parameter like this.</span></span>


```js
Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, 
    writeDataCallback);

// Callback to write the selected data to the add-in UI.
function writeDataCallback(result) {
    write('Selected data: ' + result.value);
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


## <a name="differences-in-whats-returned-to-the-asyncresultvalue-property"></a><span data-ttu-id="a9446-140">Diff?rences dans les ?l?ments retourn?s ? la propri?t? AsyncResult.value</span><span class="sxs-lookup"><span data-stu-id="a9446-140">Differences in what's returned to the AsyncResult.value property</span></span>


<span data-ttu-id="a9446-p110">Les propri?t?s  **asyncContext**,  **status** et **error** de l?objet **AsyncResult** retournent le m?me type d?informations ? la fonction de rappel pass?e ? toutes les m?thodes ? Async ?. Cependant, les ?l?ments retourn?s ? la propri?t? **AsyncResult.value** varient selon la fonctionnalit? de la m?thode ? Async ?.</span><span class="sxs-lookup"><span data-stu-id="a9446-p110">The  **asyncContext**,  **status**, and  **error** properties of the **AsyncResult** object return the same kinds of information to the callback function passed to all "Async" methods. However, what's returned to the **AsyncResult.value** property varies depending on the functionality of the "Async" method.</span></span>

<span data-ttu-id="a9446-p111">Par exemple, les m?thodes **addHandlerAsync** (des objets [Binding](https://dev.office.com/reference/add-ins/shared/binding), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart), [Document](https://dev.office.com/reference/add-ins/shared/document), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings) et [Settings](https://dev.office.com/reference/add-ins/shared/settings)) sont utilis?es pour ajouter des fonctions de gestionnaire d??v?nements aux ?l?ments repr?sent?s par ces objets. Vous pouvez acc?der ? la propri?t? **AsyncResult.value** ? partir de la fonction de rappel que vous transmettez aux m?thodes **addHandlerAsync**, mais comme vous n?acc?dez ? aucune donn?e ni ? aucun objet lorsque vous ajoutez un gestionnaire d??v?nements, la propri?t? **value** renvoie toujours **undefined** si vous tentez d?y acc?der.</span><span class="sxs-lookup"><span data-stu-id="a9446-p111">For example, the  **addHandlerAsync** methods (of the [Binding](https://dev.office.com/reference/add-ins/shared/binding), [CustomXmlPart](https://dev.office.com/reference/add-ins/shared/customxmlpart.customxmlpart), [Document](https://dev.office.com/reference/add-ins/shared/document), [RoamingSettings](https://dev.office.com/reference/add-ins/outlook/RoamingSettings), and [Settings](https://dev.office.com/reference/add-ins/shared/settings) objects) are used to add event handler functions to the items represented by these objects. You can access the **AsyncResult.value** property from the callback function you pass to any of the **addHandlerAsync** methods, but since no data or object is being accessed when you add an event handler, the **value** property always returns **undefined** if you attempt to access it.</span></span>

<span data-ttu-id="a9446-p112">En revanche, si vous appelez la m?thode  **Document.getSelectedDataAsync**, celle-ci renvoie les donn?es que l?utilisateur a s?lectionn?es dans le document ? la propri?t?  **AsyncResult.value** dans le rappel. Ou alors, si vous appelez la m?thode [Bindings.getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync), celle-ci renvoie un tableau de tous les objets  **Binding** du document. Enfin, si vous appelez la m?thode [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync), celle-ci renvoie un seul objet  **Binding**.</span><span class="sxs-lookup"><span data-stu-id="a9446-p112">On the other hand, if you call the  **Document.getSelectedDataAsync** method, it returns the data the user selected in the document to the **AsyncResult.value** property in the callback. Or, if you call the [Bindings.getAllAsync](https://dev.office.com/reference/add-ins/shared/bindings.getallasync) method, it returns an array of all of the **Binding** objects in the document. And, if you call the [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) method, it returns a single **Binding** object.</span></span>

<span data-ttu-id="a9446-p113">Pour obtenir une description des ?l?ments renvoy?s ? la propri?t? **AsyncResult.value** pour une m?thode ? Async ?, voir la section relative ? la valeur de rappel dans la rubrique de r?f?rence de cette m?thode. Pour obtenir un r?sum? de tous les objets qui fournissent des m?thodes ? Async ?, voir le tableau situ? au bas de la rubrique relative ? l?objet [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult).</span><span class="sxs-lookup"><span data-stu-id="a9446-p113">For a description of what's returned to the  **AsyncResult.value** property for an "Async" method, see the "Callback value" section of that method's reference topic. For a summary of all of the objects that provide "Async" methods, see the table at the bottom of the [AsyncResult](https://dev.office.com/reference/add-ins/shared/asyncresult) object topic.</span></span>


## <a name="asynchronous-programming-patterns"></a><span data-ttu-id="a9446-150">Mod?les de programmation asynchrone</span><span class="sxs-lookup"><span data-stu-id="a9446-150">Asynchronous programming patterns</span></span>


<span data-ttu-id="a9446-151">L?interface API JavaScript pour Office prend en charge deux types de mod?les de programmation asynchrone :</span><span class="sxs-lookup"><span data-stu-id="a9446-151">The JavaScript API for Office supports two kinds of asynchronous programming patterns:</span></span>


- <span data-ttu-id="a9446-152">Utilisation des rappels imbriqu?s</span><span class="sxs-lookup"><span data-stu-id="a9446-152">Using nested callbacks</span></span>
    
- <span data-ttu-id="a9446-153">Utilisation du mod?le des promesses</span><span class="sxs-lookup"><span data-stu-id="a9446-153">Using the promises pattern</span></span>
    
<span data-ttu-id="a9446-p114">La programmation asynchrone ? l?aide des fonctions de rappel n?cessite que vous imbriquiez fr?quemment le r?sultat retourn? d?un rappel au sein d?au moins deux rappels. Pour ce faire, vous pouvez utiliser les rappels imbriqu?s de toutes les m?thodes ? Async ? de l?API.</span><span class="sxs-lookup"><span data-stu-id="a9446-p114">Asynchronous programming with callback functions frequently requires you to nest the returned result of one callback within two or more callbacks. If you need to do so, you can use nested callbacks from all "Async" methods of the API.</span></span>

<span data-ttu-id="a9446-p115">L?utilisation des rappels imbriqu?s est un mod?le de programmation familier pour la plupart des d?veloppeurs JavaScript, mais le code contenant des rappels fortement imbriqu?s peut ?tre difficile ? lire et ? comprendre. Pour offrir une solution de remplacement aux rappels imbriqu?s, l?interface API JavaScript pour Office prend ?galement en charge l?impl?mentation du mod?le des promesses. Cependant, dans la version actuelle de l?interface API JavaScript pour Office, le mod?le des promesses fonctionne uniquement avec du code destin? aux [liaisons dans les feuilles de calcul Excel et les documents Word](bind-to-regions-in-a-document-or-spreadsheet.md).</span><span class="sxs-lookup"><span data-stu-id="a9446-p115">Using nested callbacks is a programming pattern familiar to most JavaScript developers, but code with deeply nested callbacks can be difficult to read and understand. As an alternative to nested callbacks, the JavaScript API for Office also supports an implementation of the promises pattern. However, in the current version of the JavaScript API for Office, the promises pattern only works with code for [bindings in Excel spreadsheets and Word documents](bind-to-regions-in-a-document-or-spreadsheet.md).</span></span>

<a name="AsyncProgramming_NestedCallbacks" />
### <a name="asynchronous-programming-using-nested-callback-functions"></a><span data-ttu-id="a9446-159">Programmation asynchrone utilisant des fonctions de rappel imbriqu?es</span><span class="sxs-lookup"><span data-stu-id="a9446-159">Asynchronous programming using nested callback functions</span></span>


<span data-ttu-id="a9446-p116">Vous devez fr?quemment effectuer au moins deux op?rations asynchrones pour r?aliser une t?che. Pour ce faire, vous pouvez imbriquer un appel ? Async ? dans un autre.</span><span class="sxs-lookup"><span data-stu-id="a9446-p116">Frequently, you need to perform two or more asynchronous operations to complete a task. To accomplish that, you can nest one "Async" call inside another.</span></span> 

<span data-ttu-id="a9446-162">L?exemple de code suivant imbrique deux appels asynchrones.</span><span class="sxs-lookup"><span data-stu-id="a9446-162">The following code example nests two asynchronous calls.</span></span> 


- <span data-ttu-id="a9446-p117">D?abord, la m?thode [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) est appel?e pour acc?der ? une liaison dans le document nomm? ? MyBinding ?. L?objet **AsyncResult** renvoy? au param?tre `result` de ce rappel donne acc?s ? l?objet de liaison sp?cifi? dans la propri?t? **AsyncResult.value**.</span><span class="sxs-lookup"><span data-stu-id="a9446-p117">First, the [Bindings.getByIdAsync](https://dev.office.com/reference/add-ins/shared/bindings.getbyidasync) method is called to access a binding in the document named "MyBinding". The **AsyncResult** object returned to the `result` parameter of that callback provides access to the specified binding object from the **AsyncResult.value** property.</span></span>
    
- <span data-ttu-id="a9446-165">Ensuite, l?objet Binding auquel vous avez acc?d? ? partir du premier param?tre `result` est utilis? pour appeler la m?thode [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync).</span><span class="sxs-lookup"><span data-stu-id="a9446-165">Then, the binding object accessed from the first  `result` parameter is used to call the [Binding.getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync) method.</span></span>
    
- <span data-ttu-id="a9446-166">Enfin, le param?tre  `result2` du rappel pass? ? la m?thode **Binding.getDataAsync** est utilis? pour afficher les donn?es dans la liaison.</span><span class="sxs-lookup"><span data-stu-id="a9446-166">Finally, the  `result2` parameter of the callback passed to the **Binding.getDataAsync** method is used to display the data in the binding.</span></span>
    



```js
function readData() {
    Office.context.document.bindings.getByIdAsync("MyBinding", function (result) {
        result.value.getDataAsync({ coercionType: 'text' }, function (result2) {
            write(result2.value);
        });
    });
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="a9446-167">Ce mod?le de rappel imbriqu? de base s?applique ? toutes les m?thodes asynchrones dans l?interface API JavaScript pour Office.</span><span class="sxs-lookup"><span data-stu-id="a9446-167">This basic nested callback pattern can be used for all asynchronous methods in the JavaScript API for Office.</span></span>

<span data-ttu-id="a9446-168">Les sections suivantes montrent comment utiliser des fonctions anonymes ou nomm?es pour des rappels imbriqu?s dans des m?thodes asynchrones.</span><span class="sxs-lookup"><span data-stu-id="a9446-168">The following sections show how to use either anonymous or named functions for nested callbacks in asynchronous methods.</span></span>


#### <a name="using-anonymous-functions-for-nested-callbacks"></a><span data-ttu-id="a9446-169">Utilisation des fonctions anonymes pour des rappels imbriqu?s</span><span class="sxs-lookup"><span data-stu-id="a9446-169">Using anonymous functions for nested callbacks</span></span>

<span data-ttu-id="a9446-p118">Dans l?exemple suivant, deux fonctions anonymes sont d?clar?es en ligne et pass?es dans les m?thodes  **getByIdAsync** et **getDataAsync** en tant que rappels imbriqu?s. Comme les fonctions sont tr?s simples, l?objet de l?impl?mentation est ?vident.</span><span class="sxs-lookup"><span data-stu-id="a9446-p118">In the following example, two anonymous functions are declared inline and passed into the  **getByIdAsync** and **getDataAsync** methods as nested callbacks. Because the functions are simple and inline, the intent of the implementation is immediately clear.</span></span>


```js
Office.context.document.bindings.getByIdAsync('myBinding', function (bindingResult) {
    bindingResult.value.getDataAsync(function (getResult) {
        if (getResult.status == Office.AsyncResultStatus.Failed) {
            write('Action failed. Error: ' + asyncResult.error.message);
        } else {
            write('Data has been read successfully.');
        }
    });
});

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


#### <a name="using-named-functions-for-nested-callbacks"></a><span data-ttu-id="a9446-172">Utilisation de fonctions nomm?es pour des rappels imbriqu?s</span><span class="sxs-lookup"><span data-stu-id="a9446-172">Using named functions for nested callbacks</span></span>

<span data-ttu-id="a9446-p119">Dans des impl?mentations complexes, il peut ?tre utile d?utiliser des fonctions nomm?es pour garantir une meilleure lisibilit?, simplicit? de gestion et possibilit? de r?utilisation du code. Dans l?exemple suivant, les deux fonctions anonymes de l?exemple dans la section pr?c?dente ont ?t? r??crites comme fonctions nomm?es  `deleteAllData` et `showResult`. Ces fonctions nomm?es sont ensuite pass?es dans les m?thodes  **getByIdAsync** et **deleteAllDataValuesAsync** comme rappels par nom.</span><span class="sxs-lookup"><span data-stu-id="a9446-p119">In complex implementations, it may be helpful to use named functions to make your code easier to read, maintain, and reuse. In the following example, the two anonymous functions from the example in the previous section have been rewritten as functions named  `deleteAllData` and `showResult`. These named functions are then passed into the  **getByIdAsync** and **deleteAllDataValuesAsync** methods as callbacks by name.</span></span>


```js
Office.context.document.bindings.getByIdAsync('myBinding', deleteAllData);

function deleteAllData(asyncResult) {
    asyncResult.value.deleteAllDataValuesAsync(showResult);
}

function showResult(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Action failed. Error: ' + asyncResult.error.message);
    } else {
        write('Data has been deleted successfully.');
    }
}

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


### <a name="asynchronous-programming-using-the-promises-pattern-to-access-data-in-bindings"></a><span data-ttu-id="a9446-176">Programmation asynchrone en utilisant le mod?le des promesses pour acc?der aux donn?es des liaisons</span><span class="sxs-lookup"><span data-stu-id="a9446-176">Asynchronous programming using the promises pattern to access data in bindings</span></span>


<span data-ttu-id="a9446-p120">Plut?t que de transmettre une fonction de rappel et d?attendre le renvoi de la fonction pour poursuivre l?ex?cution, le motif de programmation des promesses renvoie imm?diatement un objet de promesse qui repr?sente le r?sultat souhait?. Toutefois, contrairement ? la vraie programmation synchrone, en arri?re-plan, la concr?tisation du r?sultat pr?vu est en fait diff?r?e jusqu?? ce que l?environnement d?ex?cution des compl?ments Office puisse r?aliser la demande. Un gestionnaire _onError_ est fourni pour couvrir les cas o? la demande ne peut pas ?tre remplie.</span><span class="sxs-lookup"><span data-stu-id="a9446-p120">Instead of passing a callback function and waiting for the function to return before execution continues, the promises programming pattern immediately returns a promise object that represents its intended result. However, unlike true synchronous programming, under the covers the fulfillment of the promised result is actually deferred until the Office Add-ins runtime environment can complete the request. An _onError_ handler is provided to cover situations when the request can't be fulfilled.</span></span>

<span data-ttu-id="a9446-p121">L?interface API JavaScript pour Office fournit la m?thode [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) pour prendre en charge le mod?le des promesses permettant d?utiliser des objets de liaison existants. L?objet de promesse renvoy? ? la m?thode **Office.select** prend en charge uniquement les quatre m?thodes auxquelles vous pouvez acc?der directement ? partir de l?objet [Binding](https://dev.office.com/reference/add-ins/shared/binding) : [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync), [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) et [removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync).</span><span class="sxs-lookup"><span data-stu-id="a9446-p121">The JavaScript API for Office provides the [Office.select](https://dev.office.com/reference/add-ins/shared/office.select) method to support the promises pattern for working with existing binding objects. The promise object returned to the **Office.select** method supports only the four methods that you can access directly from the [Binding](https://dev.office.com/reference/add-ins/shared/binding) object: [getDataAsync](https://dev.office.com/reference/add-ins/shared/binding.getdataasync), [setDataAsync](https://dev.office.com/reference/add-ins/shared/binding.setdataasync), [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value), and [removeHandlerAsync](https://dev.office.com/reference/add-ins/shared/binding.removehandlerasync).</span></span>

<span data-ttu-id="a9446-182">Le mod?le des promesses ? utiliser avec les liaisons se pr?sente comme suit :</span><span class="sxs-lookup"><span data-stu-id="a9446-182">The promises pattern for working with bindings takes this form:</span></span>

 <span data-ttu-id="a9446-183">**Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_</span><span class="sxs-lookup"><span data-stu-id="a9446-183">**Office.select(**_selectorExpression_,  _onError_**).**_BindingObjectAsyncMethod_</span></span>

<span data-ttu-id="a9446-p122">Le param?tre  _selectorExpression_ a le format `"bindings#bindingId"`, o?  _bindingId_ est le nom ( **id**) d?une liaison cr??e pr?c?demment dans le document ou la feuille de calcul (? l?aide de l?une des m?thodes ? addFrom ? de la collection  **Bindings** :  **addFromNamedItemAsync**,  **addFromPromptAsync** ou **addFromSelectionAsync**). Par exemple, l?expression de s?lecteur  `bindings#cities` sp?cifie que vous voulez acc?der ? la liaison avec le param?tre **id** 'cities'.</span><span class="sxs-lookup"><span data-stu-id="a9446-p122">The  _selectorExpression_ parameter takes the form `"bindings#bindingId"`, where  _bindingId_ is the name ( **id**) of a binding that you created previously in the document or spreadsheet (using one of the "addFrom" methods of the  **Bindings** collection: **addFromNamedItemAsync**,  **addFromPromptAsync**, or  **addFromSelectionAsync**). For example, the selector expression  `bindings#cities` specifies that you want to access the binding with an **id** of 'cities'.</span></span>

<span data-ttu-id="a9446-p123">Le param?tre  _onError_ est une fonction de gestion des erreurs qui prend un seul param?tre de type **AsyncResult** pouvant ?tre utilis? pour acc?der ? un objet **Error** si la m?thode **select** ne permet pas d?acc?der ? la liaison sp?cifi?e. L?exemple suivant montre une fonction de gestion des erreurs de base pouvant ?tre pass?e au param?tre _onError_.</span><span class="sxs-lookup"><span data-stu-id="a9446-p123">The  _onError_ parameter is an error handling function which takes a single parameter of type **AsyncResult** that can be used to access an **Error** object, if the **select** method fails to access the specified binding. The following example shows a basic error handler function that can be passed to the _onError_ parameter.</span></span>




```js
function onError(result){
    var err = result.error;
    write(err.name + ": " + err.message);
}
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="a9446-p124">Remplacez l?espace r?serv? _BindingObjectAsyncMethod_ par un appel ? l?une des quatre m?thodes d?objet **Binding** prises en charge par l?objet de promesse : **getDataAsync**, **setDataAsync**, **addHandlerAsync** ou **removeHandlerAsync**. Les appels ? ces m?thodes ne prennent pas en charge les promesses suppl?mentaires. Vous devez les appeler ? l?aide du [mod?le de fonction de rappel imbriqu?e](#AsyncProgramming_NestedCallbacks).</span><span class="sxs-lookup"><span data-stu-id="a9446-p124">Replace the  _BindingObjectAsyncMethod_ placeholder with a call to any of the four **Binding** object methods supported by the promise object: **getDataAsync**,  **setDataAsync**,  **addHandlerAsync**, or  **removeHandlerAsync**. Calls to these methods don't support additional promises. You must call them using the [nested callback function pattern](#AsyncProgramming_NestedCallbacks).</span></span>

<span data-ttu-id="a9446-p125">Une fois qu?une promesse d?objet  **Binding** est concr?tis?e, elle peut ?tre r?utilis?e dans l?appel de m?thode cha?n? comme s?il s?agissait d?une liaison (le runtime de compl?ment ne retentera pas de concr?tiser la promesse de fa?on asynchrone). Si la promesse d?objet **Binding** ne peut pas ?tre concr?tis?e, le runtime de compl?ment retentera d?acc?der ? l?objet de liaison au prochain appel de l?une de ses m?thodes asynchrones.</span><span class="sxs-lookup"><span data-stu-id="a9446-p125">After a  **Binding** object promise is fulfilled, it can be reused in the chained method call as if it were a binding (the add-in runtime won't asynchronously retry fulfilling the promise). If the **Binding** object promise can't be fulfilled, the add-in runtime will try again to access the binding object the next time one of its asynchronous methods is invoked.</span></span>

<span data-ttu-id="a9446-193">L?exemple de code suivant utilise la m?thode **select** pour r?cup?rer une liaison avec l?**id** ? `cities` ? ? partir de la collection **Bindings**, puis appelle la m?thode [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) afin d?ajouter un gestionnaire d??v?nements pour l??v?nement [dataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) de la liaison.</span><span class="sxs-lookup"><span data-stu-id="a9446-193">The following code example uses the  **select** method to retrieve a binding with the **id** " `cities`" from the  **Bindings** collection, and then calls the [addHandlerAsync](https://dev.office.com/reference/add-ins/shared/asyncresult.value) method to add an event handler for the [dataChanged](https://dev.office.com/reference/add-ins/shared/binding.bindingdatachangedevent) event of the binding.</span></span>




```js
function addBindingDataChangedEventHandler() {
    Office.select("bindings#cities", function onError(){/* error handling code */}).addHandlerAsync(Office.EventType.BindingDataChanged,
    function (eventArgs) {
        doSomethingWithBinding(eventArgs.binding);
    });
}

```


> [!IMPORTANT]
> <span data-ttu-id="a9446-p126">La promesse d?objet **Binding** renvoy?e par la m?thode **Office.select** fournit uniquement un acc?s aux quatre m?thodes de l?objet **Binding**. Pour acc?der ? l?un des autres membres de l?objet **Binding**, vous devez utiliser la propri?t? **Document.bindings** et la m?thode **Bindings.getByIdAsync** ou **Bindings.getAllAsync** pour r?cup?rer l?objet **Binding**. Par exemple, pour acc?der aux propri?t?s de l?objet **Binding** (propri?t? **document**, **id** ou **type**) ou pour acc?der aux propri?t?s de l?objet [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding) ou [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding), vous devez utiliser la m?thode **getByIdAsync** ou **getAllAsync** pour r?cup?rer un objet **Binding**.</span><span class="sxs-lookup"><span data-stu-id="a9446-p126">The  **Binding** object promise returned by the **Office.select** method provides access to only the four methods of the **Binding** object. If you need to access any of the other members of the **Binding** object, instead you must use the **Document.bindings** property and **Bindings.getByIdAsync** or **Bindings.getAllAsync** methods to retrieve the **Binding** object. For example, if you need to access any of the **Binding** object's properties (the **document**,  **id**, or  **type** properties), or need to access the properties of the [MatrixBinding](https://dev.office.com/reference/add-ins/shared/binding.matrixbinding) or [TableBinding](https://dev.office.com/reference/add-ins/shared/binding.tablebinding) objects, you must use the **getByIdAsync** or **getAllAsync** methods to retrieve a **Binding** object.</span></span>


## <a name="passing-optional-parameters-to-asynchronous-methods"></a><span data-ttu-id="a9446-197">Passage de param?tres facultatifs ? des m?thodes asynchrones</span><span class="sxs-lookup"><span data-stu-id="a9446-197">Passing optional parameters to asynchronous methods</span></span>


<span data-ttu-id="a9446-198">La syntaxe courante pour toutes les m?thodes ? Async ? suit ce mod?le :</span><span class="sxs-lookup"><span data-stu-id="a9446-198">The common syntax for all "Async" methods follows this pattern:</span></span>

 <span data-ttu-id="a9446-199">_AsyncMethod_ `(`_RequiredParameters_`, [`_OptionalParameters_`],`_CallbackFunction_ `);`</span><span class="sxs-lookup"><span data-stu-id="a9446-199">_AsyncMethod_ `(` _RequiredParameters_ `, [` _OptionalParameters_ `],` _CallbackFunction_ `);`</span></span>

<span data-ttu-id="a9446-p127">Toutes les m?thodes asynchrones prennent en charge des param?tres facultatifs, qui sont pass?s sous la forme d?un objet JSON (JavaScript Object Notation) qui contient un ou plusieurs param?tres facultatifs. L?objet JSON contenant les param?tres facultatifs est une collection non ordonn?e de paires cl?-valeur o? le caract?re ? : ? s?pare la cl? de la valeur. Chaque paire dans l?objet est s?par?e par une virgule, et l?ensemble complet de paires est plac? entre accolades. La cl? est le nom du param?tre, et la valeur est la valeur ? passer pour ce param?tre.</span><span class="sxs-lookup"><span data-stu-id="a9446-p127">All asynchronous methods support optional parameters, which are passed in as a JavaScript Object Notation (JSON) object that contains one or more optional parameters. The JSON object containing the optional parameters is an unordered collection of key-value pairs with the ":" character separating the key and the value. Each pair in the object is comma-separated, and the entire set of pairs is enclosed in braces. The key is the parameter name, and value is the value to pass for that parameter.</span></span>

<span data-ttu-id="a9446-204">Vous pouvez cr?er l?objet JSON qui contient les param?tres facultatifs incorpor?s, ou cr?er un objet  `options` et le passer comme param?tre _options_.</span><span class="sxs-lookup"><span data-stu-id="a9446-204">You can create the JSON object that contains optional parameters inline, or by creating an  `options` object and passing that in as the _options_ parameter.</span></span>


### <a name="passing-optional-parameters-inline"></a><span data-ttu-id="a9446-205">Passage de param?tres facultatifs incorpor?s</span><span class="sxs-lookup"><span data-stu-id="a9446-205">Passing optional parameters inline</span></span>

<span data-ttu-id="a9446-206">Par exemple, la syntaxe pour appeler la m?thode [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) avec des param?tres facultatifs incorpor?s se pr?sente comme ceci :</span><span class="sxs-lookup"><span data-stu-id="a9446-206">For example, the syntax for calling the [Document.setSelectedDataAsync](https://dev.office.com/reference/add-ins/shared/document.setselecteddataasync) method with optional parameters inline looks like this:</span></span>

```js
 Office.context.document.setSelectedDataAsync(data, {coercionType: 'coercionType', asyncContext:' asyncContext},callback);

```

<span data-ttu-id="a9446-207">Dans cette forme de syntaxe d?appel, les deux param?tres facultatifs,  _coercionType_ et _asyncContext_, sont d?finis comme un objet incorpor? mis entre accolades.</span><span class="sxs-lookup"><span data-stu-id="a9446-207">In this form of the calling syntax, the two optional parameters,  _coercionType_ and _asyncContext_, are defined as a JSON object inline enclosed in braces.</span></span>

<span data-ttu-id="a9446-208">L?exemple suivant montre comment appeler la m?thode **Document.setSelectedDataAsync** en sp?cifiant des param?tres facultatifs incorpor?s.</span><span class="sxs-lookup"><span data-stu-id="a9446-208">The following example shows how to call to the  **Document.setSelectedDataAsync** method by specifying optional parameters inline.</span></span>


```js
Office.context.document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    {coercionType: "html", asyncContext: 42},
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


> [!NOTE]
> <span data-ttu-id="a9446-209">Vous pouvez sp?cifier des param?tres facultatifs dans l?objet JSON dans n?importe quel ordre dans la mesure o? leurs noms sont correctement sp?cifi?s.</span><span class="sxs-lookup"><span data-stu-id="a9446-209">You can specify optional parameters in any order in the JSON object as long as their names are specified correctly.</span></span>


### <a name="passing-optional-parameters-in-an-options-object"></a><span data-ttu-id="a9446-210">Passage de param?tres facultatifs dans un objet options</span><span class="sxs-lookup"><span data-stu-id="a9446-210">Passing optional parameters in an options object</span></span>

<span data-ttu-id="a9446-211">Vous pouvez ?galement cr?er un objet nomm?  `options` qui sp?cifie les param?tres facultatifs s?par?ment de l?appel de la m?thode, puis passe l?objet `options` comme l?argument _options_.</span><span class="sxs-lookup"><span data-stu-id="a9446-211">Alternatively, you can create an object named  `options` that specifies the optional parameters separately from the method call, and then pass the `options` object as the _options_ argument.</span></span>

<span data-ttu-id="a9446-212">L?exemple suivant illustre une mani?re de cr?er l?objet  `options`, o?  `parameter1` et `value1` notamment sont des espaces r?serv?s aux noms et valeurs de param?tres effectifs.</span><span class="sxs-lookup"><span data-stu-id="a9446-212">The following example shows one way of creating the  `options` object, where `parameter1`,  `value1`, and so on, are placeholders for the actual parameter names and values.</span></span>




```js
var options = {
    parameter1: value1,
    parameter2: value2,
    ...
    parameterN: valueN
};

```

<span data-ttu-id="a9446-213">Ce qui ressemble ? l?exemple suivant lors de la sp?cification des param?tres [ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration) et [FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration).</span><span class="sxs-lookup"><span data-stu-id="a9446-213">Which looks like the following example when used to specify the [ValueFormat](https://dev.office.com/reference/add-ins/shared/valueformat-enumeration) and [FilterType](https://dev.office.com/reference/add-ins/shared/filtertype-enumeration) parameters.</span></span>




```js
var options = {
    valueFormat: "unformatted",
    filterType: "all"
};
```

<span data-ttu-id="a9446-214">Voici une autre fa?on de cr?er l?objet  `options`.</span><span class="sxs-lookup"><span data-stu-id="a9446-214">Here's another way of creating the  `options` object.</span></span>




```js
var options = {};
options[parameter1] = value1;
options[parameter2] = value2;
...
options[parameterN] = valueN;
```

<span data-ttu-id="a9446-215">Ce qui ressemble ? l?exemple suivant lors de la sp?cification des param?tres  **ValueFormat** et **FilterType** :</span><span class="sxs-lookup"><span data-stu-id="a9446-215">Which looks like the following example when used to specify the  **ValueFormat** and **FilterType** parameters.:</span></span>


```js
var options = {};
options["ValueFormat"] = "unformatted";
options["FilterType"] = "all";
```


> [!NOTE]
> <span data-ttu-id="a9446-216">Au moment de cr?er l?objet `options` en employant l?une ou l?autre de ces m?thodes, vous pouvez sp?cifier des param?tres facultatifs dans n?importe quel ordre du moment o? leurs noms sont sp?cifi?s correctement.</span><span class="sxs-lookup"><span data-stu-id="a9446-216">When using either method of creating the  `options` object, you can specify optional parameters in any order as long as their names are specified correctly.</span></span>

<span data-ttu-id="a9446-217">L?exemple suivant illustre comment appeler la m?thode **Document.setSelectedDataAsync** en sp?cifiant des param?tres facultatifs dans un objet `options`.</span><span class="sxs-lookup"><span data-stu-id="a9446-217">The following example shows how to call to the  **Document.setSelectedDataAsync** method by specifying optional parameters in an `options` object.</span></span>




```js
var options = {
   coercionType: "html",
   asyncContext: 42
};

document.setSelectedDataAsync(
    "<html><body>hello world</body></html>",
    options,
    function(asyncResult) {
        write(asyncResult.status + " " + asyncResult.asyncContext);
    }
)

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```


<span data-ttu-id="a9446-p128">Dans les deux exemples de param?tres facultatifs, le param?tre _callback_ est sp?cifi? comme le dernier param?tre (? la suite des param?tres facultatifs incorpor?s, ou de l?objet de l?argument _options_). Vous pouvez ?galement sp?cifier le param?tre _callback_ ? l?int?rieur de l?objet JSON incorpor?, ou dans l?objet `options`. Cependant, vous ne pouvez passer le param?tre _callback_ qu?? un seul endroit : soit dans l?objet _options_ (incorpor? ou cr?? en externe), soit comme dernier param?tre, mais pas les deux.</span><span class="sxs-lookup"><span data-stu-id="a9446-p128">In both optional parameter examples, the  _callback_ parameter is specified as the last parameter (following the inline optional parameters, or following the _options_ argument object). Alternatively, you can specify the _callback_ parameter inside either the inline JSON object, or in the `options` object. However, you can pass the _callback_ parameter in only one location: either in the _options_ object (inline or created externally), or as the last parameter, but not both.</span></span>


## <a name="see-also"></a><span data-ttu-id="a9446-221">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="a9446-221">See also</span></span>

- [<span data-ttu-id="a9446-222">Pr?sentation de l?API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="a9446-222">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md) 
- [<span data-ttu-id="a9446-223">Interface API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="a9446-223">JavaScript API for Office</span></span>](https://dev.office.com/reference/add-ins/javascript-api-for-office)
     
