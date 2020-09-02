---
title: Utilisation du modèle d’API propre à l’application
description: Découvrez le modèle d’API basée sur la promesse pour les compléments Excel, OneNote et Word.
ms.date: 07/29/2020
localization_priority: Normal
ms.openlocfilehash: cabd1ea0076b672a1dbda3079a767b0e8a1a62b7
ms.sourcegitcommit: 4adfc368a366f00c3f3d7ed387f34aaecb47f17c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/01/2020
ms.locfileid: "47326281"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="1dded-103">Utilisation du modèle d’API propre à l’application</span><span class="sxs-lookup"><span data-stu-id="1dded-103">Using the application-specific API model</span></span>

<span data-ttu-id="1dded-104">Cet article explique comment utiliser le modèle d’API pour créer des compléments dans Excel, Word et OneNote.</span><span class="sxs-lookup"><span data-stu-id="1dded-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="1dded-105">Il présente les concepts fondamentaux de l’utilisation des API basées sur les promesses.</span><span class="sxs-lookup"><span data-stu-id="1dded-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="1dded-106">Ce modèle n’est pas pris en charge par les clients Office 2013.</span><span class="sxs-lookup"><span data-stu-id="1dded-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="1dded-107">Utilisez le [modèle d’API commun](office-javascript-api-object-model.md) pour utiliser ces versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="1dded-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="1dded-108">Pour obtenir des notes sur la disponibilité complète de la plateforme, consultez la rubrique [Office client Application and Platform Availability for Office Add-ins](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="1dded-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="1dded-109">Les exemples de cette page utilisent les API JavaScript Excel, mais les concepts s’appliquent également aux API JavaScript OneNote, Visio et Word.</span><span class="sxs-lookup"><span data-stu-id="1dded-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="1dded-110">Nature asynchrone des API basées sur les promesses</span><span class="sxs-lookup"><span data-stu-id="1dded-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="1dded-111">Les compléments Office sont des sites Web qui s’affichent dans un conteneur de navigateur dans les applications Office, telles qu’Excel.</span><span class="sxs-lookup"><span data-stu-id="1dded-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="1dded-112">Ce conteneur est incorporé dans l’application Office sur les plateformes de bureau, comme Office sur Windows, et s’exécute dans un iFrame HTML dans Office sur le Web.</span><span class="sxs-lookup"><span data-stu-id="1dded-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="1dded-113">Pour des raisons de performances, les API Office.js ne peuvent pas interagir de façon synchrone avec les applications Office sur toutes les plateformes.</span><span class="sxs-lookup"><span data-stu-id="1dded-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="1dded-114">Par conséquent, l' `sync()` appel de l’API dans Office.js renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est résolue lorsque l’application Office termine les actions de lecture ou d’écriture demandées.</span><span class="sxs-lookup"><span data-stu-id="1dded-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="1dded-115">En outre, vous pouvez mettre en file d’attente plusieurs actions, telles que définir des propriétés ou appeler des méthodes, et les exécuter sous la forme d’un lot de commandes avec un seul appel à `sync()` , au lieu d’envoyer une demande distincte pour chaque action.</span><span class="sxs-lookup"><span data-stu-id="1dded-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="1dded-116">Les sections suivantes décrivent comment effectuer cette procédure à l’aide des `run()` `sync()` API et.</span><span class="sxs-lookup"><span data-stu-id="1dded-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="1dded-117">fonction \*. Run</span><span class="sxs-lookup"><span data-stu-id="1dded-117">\*.run function</span></span>

<span data-ttu-id="1dded-118">`Excel.run`, `Word.run` et `OneNote.run` exécutent une fonction qui spécifie les actions à effectuer sur Excel, Word et OneNote.</span><span class="sxs-lookup"><span data-stu-id="1dded-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="1dded-119">`*.run` crée automatiquement un contexte de demande que vous pouvez utiliser pour interagir avec les objets Office.</span><span class="sxs-lookup"><span data-stu-id="1dded-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="1dded-120">`*.run`Une fois l’opération terminée, une promesse est résolue et tous les objets qui ont été alloués lors de l’exécution sont automatiquement publiés.</span><span class="sxs-lookup"><span data-stu-id="1dded-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="1dded-121">L’exemple suivant montre comment utiliser `Excel.run` .</span><span class="sxs-lookup"><span data-stu-id="1dded-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="1dded-122">Le même modèle est également utilisé avec Word et OneNote.</span><span class="sxs-lookup"><span data-stu-id="1dded-122">The same pattern is also used with Word and OneNote.</span></span>

```js
Excel.run(function (context) {
    // Add your Excel JS API calls here that will be batched and sent to the workbook.
    console.log('Your code goes here.');
}).catch(function (error) {
    // Catch and log any errors that occur within `Excel.run`.
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="request-context"></a><span data-ttu-id="1dded-123">Contexte de demande</span><span class="sxs-lookup"><span data-stu-id="1dded-123">Request context</span></span>

<span data-ttu-id="1dded-124">L’application Office et votre complément s’exécutent dans deux processus différents.</span><span class="sxs-lookup"><span data-stu-id="1dded-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="1dded-125">Dans la mesure où ils utilisent des environnements d’exécution différents, les compléments nécessitent un `RequestContext` objet pour connecter votre complément à des objets dans Office, tels que des feuilles de calcul, des plages, des paragraphes et des tableaux.</span><span class="sxs-lookup"><span data-stu-id="1dded-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="1dded-126">Cet `RequestContext` objet est fourni en tant qu’argument lors de l’appel `*.run` .</span><span class="sxs-lookup"><span data-stu-id="1dded-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="1dded-127">Objets de proxy</span><span class="sxs-lookup"><span data-stu-id="1dded-127">Proxy objects</span></span>

<span data-ttu-id="1dded-128">Les objets JavaScript Office que vous déclarez et utilisez avec les API basées sur les promesses sont des objets proxy.</span><span class="sxs-lookup"><span data-stu-id="1dded-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="1dded-129">Les méthodes que vous appelez ou les propriétés que vous définissez ou chargez sur les objets proxy sont simplement ajoutées à une file d’attente de commandes en attente.</span><span class="sxs-lookup"><span data-stu-id="1dded-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="1dded-130">Lorsque vous appelez la `sync()` méthode sur le contexte de la demande (par exemple, `context.sync()` ), les commandes en file d’attente sont envoyées vers l’application Office et exécutées.</span><span class="sxs-lookup"><span data-stu-id="1dded-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="1dded-131">Ces API sont fondamentalement centrées sur les lots.</span><span class="sxs-lookup"><span data-stu-id="1dded-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="1dded-132">Vous pouvez mettre en file d’attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la `sync()` méthode pour exécuter le lot de commandes en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="1dded-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="1dded-133">Par exemple, l’extrait de code suivant déclare l’objet [Excel. Range](/javascript/api/excel/excel.range) JavaScript local, `selectedRange` , pour faire référence à une plage sélectionnée dans le classeur Excel, puis définit certaines propriétés sur cet objet.</span><span class="sxs-lookup"><span data-stu-id="1dded-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="1dded-134">L' `selectedRange` objet est un objet proxy, de sorte que les propriétés qui sont définies et la méthode appelée sur cet objet ne sont pas reflétées dans le document Excel tant que votre complément n’a pas été appelé `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="1dded-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="1dded-135">Conseil de performance : réduisez le nombre d’objets proxy créés</span><span class="sxs-lookup"><span data-stu-id="1dded-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="1dded-136">Éviter de créer le même objet proxy à plusieurs reprises.</span><span class="sxs-lookup"><span data-stu-id="1dded-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="1dded-137">Au lieu de cela, si vous avez besoin du même objet proxy pour plus d’une opération, créez-le une seule fois et affectez-le à une variable, puis utilisez cette variable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="1dded-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

```js
// BAD: Repeated calls to .getRange() to create the same proxy object.
worksheet.getRange("A1").format.fill.color = "red";
worksheet.getRange("A1").numberFormat = "0.00%";
worksheet.getRange("A1").values = [[1]];

// GOOD: Create the range proxy object once and assign to a variable.
var range = worksheet.getRange("A1")
range.format.fill.color = "red";
range.numberFormat = "0.00%";
range.values = [[1]];

// ALSO GOOD: Use a "set" method to immediately set all the properties without even needing to create a variable!
worksheet.getRange("A1").set({
    numberFormat: [["0.00%"]],
    values: [[1]],
    format: {
        fill: {
            color: "red"
        }
    }
});
```

### <a name="sync"></a><span data-ttu-id="1dded-138">Sync()</span><span class="sxs-lookup"><span data-stu-id="1dded-138">sync()</span></span>

<span data-ttu-id="1dded-139">L’appel de la `sync()` méthode sur le contexte de demande synchronise l’état entre les objets proxy et les objets dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="1dded-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="1dded-140">La `sync()` méthode exécute toutes les commandes qui sont placées en file d’attente dans le contexte de la demande et récupère des valeurs pour les propriétés qui doivent être chargées sur les objets proxy.</span><span class="sxs-lookup"><span data-stu-id="1dded-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="1dded-141">La `sync()` méthode s’exécute de façon asynchrone et renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est résolue à la fin de la `sync()` méthode.</span><span class="sxs-lookup"><span data-stu-id="1dded-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="1dded-142">L’exemple suivant montre une fonction batch qui définit un objet proxy JavaScript local ( `selectedRange` ), charge une propriété de cet objet, puis utilise le modèle de promet JavaScript pour appeler `context.sync()` pour synchroniser l’état entre les objets proxy et les objets dans le document Excel.</span><span class="sxs-lookup"><span data-stu-id="1dded-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

```js
Excel.run(function (context) {
    var selectedRange = context.workbook.getSelectedRange();
    selectedRange.load('address');
    return context.sync()
      .then(function () {
        console.log('The selected range is: ' + selectedRange.address);
    });
}).catch(function (error) {
    console.log('error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

<span data-ttu-id="1dded-143">Dans l’exemple précédent, `selectedRange` est configuré et sa propriété `address` est chargée lorsque `context.sync()` est appelé.</span><span class="sxs-lookup"><span data-stu-id="1dded-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="1dded-144">Étant donné qu’il `sync()` s’agit d’une opération asynchrone, vous devez toujours retourner l' `Promise` objet pour vous assurer que l' `sync()` opération se termine avant que le script continue à s’exécuter.</span><span class="sxs-lookup"><span data-stu-id="1dded-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="1dded-145">Si vous utilisez la machine à écrire ou ES6 + JavaScript, vous pouvez `await` `context.sync()` appeler au lieu de renvoyer la promesse.</span><span class="sxs-lookup"><span data-stu-id="1dded-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="1dded-146">Conseil de performance : réduisez le nombre d’appels de synchronisation</span><span class="sxs-lookup"><span data-stu-id="1dded-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="1dded-147">Dans l’API JavaScript Excel, `sync()` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="1dded-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="1dded-148">Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez `sync()` et mettre en file d’attente autant de modifications que possible avant d’appeler.</span><span class="sxs-lookup"><span data-stu-id="1dded-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="1dded-149">Pour plus d’informations sur l’optimisation des performances avec `sync()` , reportez-vous à [la rubrique éviter d’utiliser la méthode Context. Sync dans les boucles](../concepts/correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="1dded-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="1dded-150">load()</span><span class="sxs-lookup"><span data-stu-id="1dded-150">load()</span></span>

<span data-ttu-id="1dded-151">Avant de pouvoir lire les propriétés d’un objet proxy, vous devez charger explicitement les propriétés pour remplir l’objet proxy avec les données du document Office, puis appeler `context.sync()` .</span><span class="sxs-lookup"><span data-stu-id="1dded-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="1dded-152">Par exemple, si vous créez un objet proxy pour référencer une plage sélectionnée, puis que vous souhaitez lire la propriété de la plage sélectionnée `address` , vous devez charger la `address` propriété avant de pouvoir la lire.</span><span class="sxs-lookup"><span data-stu-id="1dded-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="1dded-153">Pour demander le chargement des propriétés d’un objet proxy, appelez la `load()` méthode sur l’objet et spécifiez les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="1dded-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="1dded-154">L’exemple suivant montre la `Range.address` propriété en cours de chargement `myRange` .</span><span class="sxs-lookup"><span data-stu-id="1dded-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

```js
Excel.run(function (context) {
    var sheetName = 'Sheet1';
    var rangeAddress = 'A1:B2';
    var myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);

    myRange.load('address');

    return context.sync()
      .then(function () {
        console.log (myRange.address);   // ok
        //console.log (myRange.values);  // not ok as it was not loaded
        });
    }).then(function () {
        console.log('done');
}).catch(function (error) {
    console.log('Error: ' + error);
    if (error instanceof OfficeExtension.Error) {
        console.log('Debug info: ' + JSON.stringify(error.debugInfo));
    }
});
```

> [!NOTE]
> <span data-ttu-id="1dded-155">Si vous appelez uniquement des méthodes ou définissez des propriétés sur un objet proxy, il n’est pas nécessaire d’appeler la `load()` méthode.</span><span class="sxs-lookup"><span data-stu-id="1dded-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="1dded-156">La `load()` méthode n’est requise que si vous souhaitez lire les propriétés sur un objet proxy.</span><span class="sxs-lookup"><span data-stu-id="1dded-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="1dded-p115">À l’instar des demandes de définition de propriétés ou d’appel de méthodes sur des objets proxy, des demandes de chargement de propriétés sur des objets proxy sont ajoutées à la file d’attente des commandes sur le contexte de demande, qui s’exécutera la prochaine fois que vous appellerez la méthode `sync()`. Vous pouvez mettre en file d’attente autant d’appels `load()` sur le contexte de la demande que nécessaire.</span><span class="sxs-lookup"><span data-stu-id="1dded-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="1dded-159">Propriétés scalaires et de navigation</span><span class="sxs-lookup"><span data-stu-id="1dded-159">Scalar and navigation properties</span></span>

<span data-ttu-id="1dded-160">Il existe deux catégories de propriétés: **scalaire** et **de navigation**.</span><span class="sxs-lookup"><span data-stu-id="1dded-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="1dded-161">Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON.</span><span class="sxs-lookup"><span data-stu-id="1dded-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="1dded-162">Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont affectés au lieu d’affecter directement la propriété.</span><span class="sxs-lookup"><span data-stu-id="1dded-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="1dded-163">Par exemple, `name` les `position` membres de l’objet [Excel. Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que les `protection` Propriétés de `tables` navigation.</span><span class="sxs-lookup"><span data-stu-id="1dded-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="1dded-164">Votre complément peut utiliser les propriétés de navigation comme chemin d’accès pour charger des propriétés scalaires spécifiques.</span><span class="sxs-lookup"><span data-stu-id="1dded-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="1dded-165">Le code suivant met en file d’attente une `load` commande pour le nom de la police utilisée par un `Excel.Range` objet, sans charger aucune autre information.</span><span class="sxs-lookup"><span data-stu-id="1dded-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="1dded-166">Vous pouvez également définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès.</span><span class="sxs-lookup"><span data-stu-id="1dded-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="1dded-167">Par exemple, vous pouvez définir la taille de la police pour un `Excel.Range` à l’aide de `someRange.format.font.size = 10;` .</span><span class="sxs-lookup"><span data-stu-id="1dded-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="1dded-168">Vous n’avez pas besoin de charger la propriété avant de la définir.</span><span class="sxs-lookup"><span data-stu-id="1dded-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="1dded-169">N’oubliez pas que certaines propriétés sous un objet peuvent avoir le même nom qu’un autre objet.</span><span class="sxs-lookup"><span data-stu-id="1dded-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="1dded-170">Par exemple, `format` est une propriété sous l' `Excel.Range` objet, mais `format` elle est également un objet.</span><span class="sxs-lookup"><span data-stu-id="1dded-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="1dded-171">Par conséquent, si vous effectuez un appel tel que `range.load("format")` , cela équivaut à `range.format.load()` (une instruction indésirable vide `load()` ).</span><span class="sxs-lookup"><span data-stu-id="1dded-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="1dded-172">Pour éviter cela, votre code doit uniquement charger les « nœuds feuille » dans une arborescence d’objets.</span><span class="sxs-lookup"><span data-stu-id="1dded-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="1dded-173">Appel `load` sans paramètres (non recommandé)</span><span class="sxs-lookup"><span data-stu-id="1dded-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="1dded-174">Si vous appelez la `load()` méthode sur un objet (ou une collection) sans spécifier de paramètres, toutes les propriétés scalaires de l’objet ou des objets de la collection sont chargées.</span><span class="sxs-lookup"><span data-stu-id="1dded-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="1dded-175">Le chargement des données inutiles ralentira votre complément.</span><span class="sxs-lookup"><span data-stu-id="1dded-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="1dded-176">Vous devez toujours spécifier explicitement les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="1dded-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="1dded-177">La quantité de données renvoyées par une `load`instruction sans paramètre peut dépasser les limites de taille du service.</span><span class="sxs-lookup"><span data-stu-id="1dded-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="1dded-178">Pour réduire les risques pesant sur les compléments plus anciens, certaines propriétés ne sont pas renvoyées par `load` sans en faire la demande explicite.</span><span class="sxs-lookup"><span data-stu-id="1dded-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="1dded-179">Les propriétés suivantes sont exclues des opérations de chargement suivantes :</span><span class="sxs-lookup"><span data-stu-id="1dded-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="1dded-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="1dded-180">ClientResult</span></span>

<span data-ttu-id="1dded-181">Les méthodes dans les API basées sur la promesse qui retournent des types primitifs ont un modèle similaire pour le `load` / `sync` paradigme.</span><span class="sxs-lookup"><span data-stu-id="1dded-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="1dded-182">Par exemple, `Excel.TableCollection.getCount` obtient le nombre de tableaux dans la collection.</span><span class="sxs-lookup"><span data-stu-id="1dded-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="1dded-183">`getCount` renvoie un `ClientResult<number>` , ce qui signifie que la `value` propriété dans le renvoyé [`ClientResult`](/javascript/api/office/officeextension.clientresult) est un nombre.</span><span class="sxs-lookup"><span data-stu-id="1dded-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="1dded-184">Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé.</span><span class="sxs-lookup"><span data-stu-id="1dded-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="1dded-185">Le code suivant obtient le nombre total de tables dans un classeur Excel et enregistre ce nombre dans la console.</span><span class="sxs-lookup"><span data-stu-id="1dded-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

```js
var tableCount = context.workbook.tables.getCount();

// This sync call implicitly loads tableCount.value.
// Any other ClientResult values are loaded too.
return context.sync()
    .then(function () {
        // Trying to log the value before calling sync would throw an error.
        console.log (tableCount.value);
    });
```

### <a name="set"></a><span data-ttu-id="1dded-186">set()</span><span class="sxs-lookup"><span data-stu-id="1dded-186">set()</span></span>

<span data-ttu-id="1dded-187">La définition de propriétés sur un objet avec des propriétés de navigation imbriquées peut être laborieuse.</span><span class="sxs-lookup"><span data-stu-id="1dded-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="1dded-188">En guise d’alternative à la définition de propriétés individuelles à l’aide de chemins de navigation, comme décrit ci-dessus, vous pouvez utiliser la `object.set()` méthode qui est disponible sur les objets dans les API JavaScript à promesse.</span><span class="sxs-lookup"><span data-stu-id="1dded-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="1dded-189">Grâce à cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant soit un autre objet du même type Office.js, soit un objet JavaScript avec des propriétés structurées comme celles de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="1dded-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="1dded-p124">L’exemple de code suivant définit plusieurs propriétés de mise en forme d’une plage en appelant la méthode `set()` et en transmettant un objet JavaScript avec des noms et des types de propriétés reflétant la structure des propriétés dans l’objet `Range`. Cet exemple part du principe que des données sont présentes dans la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="1dded-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

```js
Excel.run(function (ctx) {
    var sheet = ctx.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:E2");
    range.set({
        format: {
            fill: {
                color: '#4472C4'
            },
            font: {
                name: 'Verdana',
                color: 'white'
            }
        }
    });
    range.format.autofitColumns();

    return ctx.sync();
}).catch(function(error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
        console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
});
```

## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="1dded-192">&#42;des méthodes et des propriétés de OrNullObject</span><span class="sxs-lookup"><span data-stu-id="1dded-192">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="1dded-193">Certaines propriétés et méthodes d’accesseur génèrent une exception lorsque l’objet souhaité n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="1dded-193">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="1dded-194">Par exemple, si vous tentez d’obtenir une feuille de calcul Excel en spécifiant un nom de feuille de calcul qui ne se trouve pas dans le classeur, la `getItem()` méthode génère une `ItemNotFound` exception.</span><span class="sxs-lookup"><span data-stu-id="1dded-194">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="1dded-195">Les bibliothèques spécifiques de l’application permettent à votre code de tester l’existence d’entités de document sans nécessiter de code de gestion des exceptions.</span><span class="sxs-lookup"><span data-stu-id="1dded-195">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="1dded-196">Pour ce faire, vous utilisez les `*OrNullObject` variantes des méthodes et des propriétés.</span><span class="sxs-lookup"><span data-stu-id="1dded-196">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="1dded-197">Ces variantes renvoient un objet dont `isNullObject` la propriété est définie sur `true` , si l’élément spécifié n’existe pas, au lieu de lever une exception.</span><span class="sxs-lookup"><span data-stu-id="1dded-197">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="1dded-198">Par exemple, vous pouvez appeler la `getItemOrNullObject()` méthode sur une collection telle que **Worksheets** pour récupérer un élément de la collection.</span><span class="sxs-lookup"><span data-stu-id="1dded-198">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="1dded-199">La `getItemOrNullObject()` méthode renvoie l’élément spécifié s’il existe ; sinon, elle renvoie un objet dont la `isNullObject` propriété a la valeur `true` .</span><span class="sxs-lookup"><span data-stu-id="1dded-199">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="1dded-200">Votre code peut ensuite évaluer cette propriété pour déterminer si l’objet existe.</span><span class="sxs-lookup"><span data-stu-id="1dded-200">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="1dded-201">Les `*OrNullObject` variantes ne renvoient jamais la valeur JavaScript `null` .</span><span class="sxs-lookup"><span data-stu-id="1dded-201">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="1dded-202">Elles renvoient des objets proxy Office ordinaires.</span><span class="sxs-lookup"><span data-stu-id="1dded-202">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="1dded-203">Si l’entité représentée par l’objet n’existe pas, la `isNullObject` propriété de l’objet est définie sur `true` .</span><span class="sxs-lookup"><span data-stu-id="1dded-203">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="1dded-204">Ne Testez pas l’objet renvoyé pour null ou falsity.</span><span class="sxs-lookup"><span data-stu-id="1dded-204">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="1dded-205">Il ne s’agit jamais `null` , `false` ou `undefined` .</span><span class="sxs-lookup"><span data-stu-id="1dded-205">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="1dded-206">L’exemple de code suivant tente de récupérer une feuille de calcul Excel nommée « Data » à l’aide de la `getItemOrNullObject()` méthode.</span><span class="sxs-lookup"><span data-stu-id="1dded-206">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="1dded-207">S’il n’existe pas de feuille de calcul portant ce nom, une nouvelle feuille est créée.</span><span class="sxs-lookup"><span data-stu-id="1dded-207">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="1dded-208">Notez que le code ne charge pas la `isNullObject` propriété.</span><span class="sxs-lookup"><span data-stu-id="1dded-208">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="1dded-209">Office charge automatiquement cette propriété lorsque `context.sync` est appelé, de sorte que vous n’avez pas besoin de le charger explicitement avec des éléments tels que `datasheet.load('isNullObject')` .</span><span class="sxs-lookup"><span data-stu-id="1dded-209">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

```js
var dataSheet = context.workbook.worksheets.getItemOrNullObject("Data");

return context.sync()
    .then(function () {
        if (dataSheet.isNullObject) {
            dataSheet = context.workbook.worksheets.add("Data");
        }

        // Set `dataSheet` to be the second worksheet in the workbook.
        dataSheet.position = 1;
    });
```

## <a name="see-also"></a><span data-ttu-id="1dded-210">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="1dded-210">See also</span></span>

* [<span data-ttu-id="1dded-211">Modèle d’objet d’API JavaScript courant</span><span class="sxs-lookup"><span data-stu-id="1dded-211">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* <span data-ttu-id="1dded-212">[Problèmes courants liés au code et comportements de plateforme inattendus](common-coding-issues.md).</span><span class="sxs-lookup"><span data-stu-id="1dded-212">[Common coding issues and unexpected platform behaviors](common-coding-issues.md).</span></span>
* [<span data-ttu-id="1dded-213">Limites des ressources et optimisation des performances pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="1dded-213">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
