---
title: Utilisation du modèle de l’API propre à l’application
description: Découvrez le modèle d’API basé sur la promesse pour les compléments Excel, OneNote et Word.
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: fb25201174dcd97b40ccf6be69b238951103db07
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408599"
---
# <a name="using-the-application-specific-api-model"></a><span data-ttu-id="0152d-103">Utilisation du modèle de l’API propre à l’application</span><span class="sxs-lookup"><span data-stu-id="0152d-103">Using the application-specific API model</span></span>

<span data-ttu-id="0152d-104">Cet article décrit l’utilisation du modèle d’API pour la création de compléments dans Excel, Word et OneNote.</span><span class="sxs-lookup"><span data-stu-id="0152d-104">This article describes how to use the API model for building add-ins in Excel, Word, and OneNote.</span></span> <span data-ttu-id="0152d-105">Il présente les concepts fondamentaux de l’utilisation des API basées sur la promesse.</span><span class="sxs-lookup"><span data-stu-id="0152d-105">It introduces core concepts that are fundamental to using the promise-based APIs.</span></span>

> [!NOTE]
> <span data-ttu-id="0152d-106">Ce modèle n’est pas pris en charge par les clients Office 2013.</span><span class="sxs-lookup"><span data-stu-id="0152d-106">This model is not supported by Office 2013 clients.</span></span> <span data-ttu-id="0152d-107">Utilisez les [Modèles communs de l’API](office-javascript-api-object-model.md) pour fonctionner avec ces versions d’Office.</span><span class="sxs-lookup"><span data-stu-id="0152d-107">Use the [Common API model](office-javascript-api-object-model.md) to work with those Office versions.</span></span> <span data-ttu-id="0152d-108">Pour consulter les notes sur la disponibilité complète des plateformes, consultez les [disponibilités de l’application et de la plateforme cliente Office pour les compléments Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="0152d-108">For full platform availability notes, see [Office client application and platform availability for Office Add-ins](../overview/office-add-in-availability.md).</span></span>

> [!TIP]
> <span data-ttu-id="0152d-109">Les exemples de cette page utilisent les API JavaScript Excel, mais les concepts s’appliquent également aux API JavaScript OneNote, Visio et Word.</span><span class="sxs-lookup"><span data-stu-id="0152d-109">The examples in this page use the Excel JavaScript APIs, but the concepts also apply to OneNote, Visio, and Word JavaScript APIs.</span></span>

## <a name="asynchronous-nature-of-the-promise-based-apis"></a><span data-ttu-id="0152d-110">Nature asynchrone des API basées sur la promesse</span><span class="sxs-lookup"><span data-stu-id="0152d-110">Asynchronous nature of the promise-based APIs</span></span>

<span data-ttu-id="0152d-111">Les compléments Office sont des sites web qui apparaissent à l’intérieur d’un conteneur de navigateur au sein des applications Office, telles qu’Excel.</span><span class="sxs-lookup"><span data-stu-id="0152d-111">Office Add-ins are websites which appear inside a browser container within Office applications, such as Excel.</span></span> <span data-ttu-id="0152d-112">Ce conteneur est incorporé dans l’application Office sur des plateformes de bureau, telles qu’Office sur Windows, et s’exécute dans un iFrame HTML, dans Office pour le web.</span><span class="sxs-lookup"><span data-stu-id="0152d-112">This container is embedded within the Office application on desktop-based platforms, such as Office on Windows, and runs inside an HTML iFrame in Office on the web.</span></span> <span data-ttu-id="0152d-113">En raison de considérations en relation avec les performances, les API Office.js ne peuvent pas interagir de façon synchronisée avec les applications Office sur toutes les plateformes.</span><span class="sxs-lookup"><span data-stu-id="0152d-113">Due to performance considerations, the Office.js APIs cannot interact synchronously with the Office applications across all platforms.</span></span> <span data-ttu-id="0152d-114">Par conséquent, l’appel de l’API `sync()` dans Office.js renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est résolue lorsque l’application Excel termine les actions de lecture ou d’écriture demandées.</span><span class="sxs-lookup"><span data-stu-id="0152d-114">Therefore, the `sync()` API call in Office.js returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Office application completes the requested read or write actions.</span></span> <span data-ttu-id="0152d-115">En outre, vous pouvez mettre en file d’attente plusieurs actions, comme la définition des propriétés ou l’appel de méthodes, et les exécuter en tant que lot de commandes avec un seul appel à `sync()`, au lieu d’envoyer une demande distincte pour chaque action.</span><span class="sxs-lookup"><span data-stu-id="0152d-115">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to `sync()`, rather than sending a separate request for each action.</span></span> <span data-ttu-id="0152d-116">Les sections suivantes décrivent comment effectuer cette tâche à l’aide des API `run()` et `sync()`.</span><span class="sxs-lookup"><span data-stu-id="0152d-116">The following sections describe how to accomplish this using the `run()` and `sync()` APIs.</span></span>

## <a name="run-function"></a><span data-ttu-id="0152d-117">\*fonction .run</span><span class="sxs-lookup"><span data-stu-id="0152d-117">\*.run function</span></span>

<span data-ttu-id="0152d-118">`Excel.run`, `Word.run`et `OneNote.run` exécutent une fonction qui spécifie les actions à effectuer dans Excel, Word et OneNote.</span><span class="sxs-lookup"><span data-stu-id="0152d-118">`Excel.run`, `Word.run`, and `OneNote.run` execute a function that specifies the actions to perform against Excel, Word, and OneNote.</span></span> <span data-ttu-id="0152d-119">`*.run` crée automatiquement un contexte de demande que vous pouvez utiliser pour interagir avec des objets Office.</span><span class="sxs-lookup"><span data-stu-id="0152d-119">`*.run` automatically creates a request context that you can use to interact with Office objects.</span></span> <span data-ttu-id="0152d-120">Lorsque `*.run` a terminé, une promesse est résolue et tous les objets alloués lors de l’exécution sont automatiquement publiés.</span><span class="sxs-lookup"><span data-stu-id="0152d-120">When `*.run` completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>

<span data-ttu-id="0152d-121">L’exemple suivant vous montre comment utiliser `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="0152d-121">The following example shows how to use `Excel.run`.</span></span> <span data-ttu-id="0152d-122">Le même modèle est également utilisé avec Word et OneNote.</span><span class="sxs-lookup"><span data-stu-id="0152d-122">The same pattern is also used with Word and OneNote.</span></span>

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

## <a name="request-context"></a><span data-ttu-id="0152d-123">Contexte de demande</span><span class="sxs-lookup"><span data-stu-id="0152d-123">Request context</span></span>

<span data-ttu-id="0152d-124">L’application Office et votre complément s’exécutent selon deux processus différents.</span><span class="sxs-lookup"><span data-stu-id="0152d-124">The Office application and your add-in run in two different processes.</span></span> <span data-ttu-id="0152d-125">Dans la mesure où ils utilisent différents environnements d’exécution différents, les compléments nécessitent un objet `RequestContext` pour connecter votre complément à des objets dans Office tels que des feuilles de calcul, des plages, des paragraphes et des tableaux.</span><span class="sxs-lookup"><span data-stu-id="0152d-125">Since they use different runtime environments, add-ins require a `RequestContext` object in order to connect your add-in to objects in Office such as worksheets, ranges, paragraphs, and tables.</span></span> <span data-ttu-id="0152d-126">Cet objet `RequestContext` est fourni en tant qu’argument lors de l’appel de `*.run`.</span><span class="sxs-lookup"><span data-stu-id="0152d-126">This `RequestContext` object is provided as an argument when calling `*.run`.</span></span>

## <a name="proxy-objects"></a><span data-ttu-id="0152d-127">Objets proxy</span><span class="sxs-lookup"><span data-stu-id="0152d-127">Proxy objects</span></span>

<span data-ttu-id="0152d-128">Les objets JavaScript Office que vous déclarez et utilisez avec les API basées sur la promesse sont des objets proxy.</span><span class="sxs-lookup"><span data-stu-id="0152d-128">The Office JavaScript objects that you declare and use with the promise-based APIs are proxy objects.</span></span> <span data-ttu-id="0152d-129">Les méthodes que vous appelez ou les propriétés que vous définissez ou chargez sur les objets proxy sont simplement ajoutées à une file d’attente de commandes en attente.</span><span class="sxs-lookup"><span data-stu-id="0152d-129">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="0152d-130">Lorsque vous appelez la méthode `sync()` dans le contexte de la demande (par exemple, `context.sync()`), les commandes en attente sont envoyées à l’application Office et s’exécutent.</span><span class="sxs-lookup"><span data-stu-id="0152d-130">When you call the `sync()` method on the request context (for example, `context.sync()`), the queued commands are dispatched to the Office application and run.</span></span> <span data-ttu-id="0152d-131">Ces API sont essentiellement centrées sur les lots.</span><span class="sxs-lookup"><span data-stu-id="0152d-131">These APIs are fundamentally batch-centric.</span></span> <span data-ttu-id="0152d-132">Vous pouvez mettre en file d’attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la méthode `sync()` pour exécuter le lot de commandes mises en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="0152d-132">You can queue up as many changes as you wish on the request context, and then call the `sync()` method to run the batch of queued commands.</span></span>

<span data-ttu-id="0152d-133">Par exemple, l’extrait de code suivant déclare l’objet JavaScript [Excel.Range](/javascript/api/excel/excel.range), `selectedRange`, pour référencer une plage sélectionnée dans la feuille de calcul Excel, et définit certaines propriétés sur cet objet.</span><span class="sxs-lookup"><span data-stu-id="0152d-133">For example, the following code snippet declares the local JavaScript [Excel.Range](/javascript/api/excel/excel.range) object, `selectedRange`, to reference a selected range in the Excel workbook, and then sets some properties on that object.</span></span> <span data-ttu-id="0152d-134">L’objet `selectedRange` est un objet proxy. Les propriétés définies et la méthode appelée sur cet objet ne seront pas répercutées dans le document Excel tant que votre complément n’a pas appelé `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="0152d-134">The `selectedRange` object is a proxy object, so the properties that are set and the method that is invoked on that object will not be reflected in the Excel document until your add-in calls `context.sync()`.</span></span>

```js
var selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```

### <a name="performance-tip-minimize-the-number-of-proxy-objects-created"></a><span data-ttu-id="0152d-135">Conseil de performance : réduire le nombre d’objets proxy créés</span><span class="sxs-lookup"><span data-stu-id="0152d-135">Performance tip: Minimize the number of proxy objects created</span></span>

<span data-ttu-id="0152d-136">Éviter de créer le même objet proxy à plusieurs reprises.</span><span class="sxs-lookup"><span data-stu-id="0152d-136">Avoid repeatedly creating the same proxy object.</span></span> <span data-ttu-id="0152d-137">Au lieu de cela, si vous avez besoin du même objet proxy pour plus d’une opération, créez-le une seule fois et affectez-le à une variable, puis utilisez cette variable dans votre code.</span><span class="sxs-lookup"><span data-stu-id="0152d-137">Instead, if you need the same proxy object for more than one operation, create it once and assign it to a variable, then use that variable in your code.</span></span>

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

### <a name="sync"></a><span data-ttu-id="0152d-138">sync()</span><span class="sxs-lookup"><span data-stu-id="0152d-138">sync()</span></span>

<span data-ttu-id="0152d-139">La méthode `sync()` concernant le contexte de demande synchronise l’état entre des objets proxy et des objets dans le document Office.</span><span class="sxs-lookup"><span data-stu-id="0152d-139">Calling the `sync()` method on the request context synchronizes the state between proxy objects and objects in the Office document.</span></span> <span data-ttu-id="0152d-140">La méthode `sync()` exécute les commandes mises en file d’attente concernant le contexte de demande et récupère des valeurs pour les propriétés qui doivent être chargées dans les objets proxy.</span><span class="sxs-lookup"><span data-stu-id="0152d-140">The `sync()` method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="0152d-141">La méthode `sync()` est exécutée de façon asynchrone et renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est résolue lorsque la méthode `sync()` est terminée.</span><span class="sxs-lookup"><span data-stu-id="0152d-141">The `sync()` method executes asynchronously and returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the `sync()` method completes.</span></span>

<span data-ttu-id="0152d-142">L’exemple suivant montre une fonction de traitement par lot qui définit un objet proxy JavaScript local (`selectedRange`), charge une propriété de cet objet et utilise ensuite le modèle de promesses JavaScript pour appeler `context.sync()` afin de synchroniser l’état entre les objets proxy et les objets du document Excel.</span><span class="sxs-lookup"><span data-stu-id="0152d-142">The following example shows a batch function that defines a local JavaScript proxy object (`selectedRange`), loads a property of that object, and then uses the JavaScript promises pattern to call `context.sync()` to synchronize the state between proxy objects and objects in the Excel document.</span></span>

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

<span data-ttu-id="0152d-143">Dans l’exemple précédent, `selectedRange` est configuré et sa propriété `address` est chargée lorsque `context.sync()` est appelé.</span><span class="sxs-lookup"><span data-stu-id="0152d-143">In the previous example, `selectedRange` is set and its `address` property is loaded when `context.sync()` is called.</span></span>

<span data-ttu-id="0152d-144">Étant donné que `sync()` est une opération asynchrone, vous devez toujours renvoyer l’objet `Promise` pour vous assurer que l’opération de `sync()` se termine avant que le script continue à s’exécuter.</span><span class="sxs-lookup"><span data-stu-id="0152d-144">Since `sync()` is an asynchronous operation, you should always return the `Promise` object to ensure the `sync()` operation completes before the script continues to run.</span></span> <span data-ttu-id="0152d-145">Si vous utilisez TypeScript ou ES6+ JavaScript, vous pouvez `await` l’appel `context.sync()` au lieu de renvoyer la promesse.</span><span class="sxs-lookup"><span data-stu-id="0152d-145">If you're using TypeScript or ES6+ JavaScript, you can `await` the `context.sync()` call instead of returning the promise.</span></span>

#### <a name="performance-tip-minimize-the-number-of-sync-calls"></a><span data-ttu-id="0152d-146">Conseil de performance : réduire le nombre d’appels de synchronisation</span><span class="sxs-lookup"><span data-stu-id="0152d-146">Performance tip: Minimize the number of sync calls</span></span>

<span data-ttu-id="0152d-147">Dans l’API JavaScript Excel, `sync()` est la seule opération asynchrone et elle peut être lente dans certaines circonstances, en particulier pour Excel sur le web.</span><span class="sxs-lookup"><span data-stu-id="0152d-147">In the Excel JavaScript API, `sync()` is the only asynchronous operation, and it can be slow under some circumstances, especially for Excel on the web.</span></span> <span data-ttu-id="0152d-148">Pour optimiser les performances, vous devez limiter le nombre de fois que vous appelez `sync()` et mettre en file d’attente autant de modifications que possible avant d’appeler.</span><span class="sxs-lookup"><span data-stu-id="0152d-148">To optimize performance, minimize the number of calls to `sync()` by queueing up as many changes as possible before calling it.</span></span> <span data-ttu-id="0152d-149">Pour plus d’informations sur l’optimisation des performances avec `sync()`, consultez [Évitez d’utiliser la méthode context.sync dans des boucles](../concepts/correlated-objects-pattern.md).</span><span class="sxs-lookup"><span data-stu-id="0152d-149">For more information about optimizing performance with `sync()`, see [Avoid using the context.sync method in loops](../concepts/correlated-objects-pattern.md).</span></span>

### <a name="load"></a><span data-ttu-id="0152d-150">load()</span><span class="sxs-lookup"><span data-stu-id="0152d-150">load()</span></span>

<span data-ttu-id="0152d-151">Pour pouvoir lire les propriétés d’un objet proxy, vous devez charger explicitement les propriétés pour remplir l’objet proxy avec les données du document Office, puis effectuer l’appel `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="0152d-151">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Office document, and then call `context.sync()`.</span></span> <span data-ttu-id="0152d-152">Par exemple, si vous créez un objet proxy pour référencer une plage sélectionnée, puis que vous voulez lire la propriété `address` de la plage sélectionnée, vous devez charger la propriété `address` avant de la lire.</span><span class="sxs-lookup"><span data-stu-id="0152d-152">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's `address` property, you need to load the `address` property before you can read it.</span></span> <span data-ttu-id="0152d-153">Pour demander le chargement des propriétés d’un objet proxy, appelez la méthode `load()` de l’objet et spécifiez les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="0152d-153">To request properties of a proxy object be loaded, call the `load()` method on the object and specify the properties to load.</span></span> <span data-ttu-id="0152d-154">L’exemple suivant illustre la propriété `Range.address` chargée pour `myRange`.</span><span class="sxs-lookup"><span data-stu-id="0152d-154">The following example shows the `Range.address` property being loaded for `myRange`.</span></span>

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
> <span data-ttu-id="0152d-155">Si vous effectuez uniquement un appel de méthodes ou que vous avez des propriétés sur un objet proxy, vous n’avez pas besoin d’effectuer l’appel de la méthode `load()`.</span><span class="sxs-lookup"><span data-stu-id="0152d-155">If you are only calling methods or setting properties on a proxy object, you don't need to call the `load()` method.</span></span> <span data-ttu-id="0152d-156">La méthode `load()` n’est requise que lorsque vous souhaitez lire les propriétés d’un objet proxy.</span><span class="sxs-lookup"><span data-stu-id="0152d-156">The `load()` method is only required when you want to read properties on a proxy object.</span></span>

<span data-ttu-id="0152d-p115">À l’instar des demandes de définition de propriétés ou d’appel de méthodes sur des objets proxy, des demandes de chargement de propriétés sur des objets proxy sont ajoutées à la file d’attente des commandes sur le contexte de demande, qui s’exécutera la prochaine fois que vous appellerez la méthode `sync()`. Vous pouvez mettre en file d’attente autant d’appels `load()` sur le contexte de la demande que nécessaire.</span><span class="sxs-lookup"><span data-stu-id="0152d-p115">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the `sync()` method. You can queue up as many `load()` calls on the request context as necessary.</span></span>

#### <a name="scalar-and-navigation-properties"></a><span data-ttu-id="0152d-159">Propriétés scalaires et de navigation</span><span class="sxs-lookup"><span data-stu-id="0152d-159">Scalar and navigation properties</span></span>

<span data-ttu-id="0152d-160">Il existe deux catégories de propriétés: **scalaire** et **de navigation**.</span><span class="sxs-lookup"><span data-stu-id="0152d-160">There are two categories of properties: **scalar** and **navigational**.</span></span> <span data-ttu-id="0152d-161">Les propriétés scalaires peuvent se voir attribuer des types, tels que des chaînes, des nombres entiers et des structures JSON.</span><span class="sxs-lookup"><span data-stu-id="0152d-161">Scalar properties are assignable types such as strings, integers, and JSON structs.</span></span> <span data-ttu-id="0152d-162">Les propriétés de navigation sont des objets en lecture seule et des collections d’objets dont les champs sont affectés, au lieu d’affecter directement la propriété.</span><span class="sxs-lookup"><span data-stu-id="0152d-162">Navigation properties are read-only objects and collections of objects that have their fields assigned, instead of directly assigning the property.</span></span> <span data-ttu-id="0152d-163">Par exemple, les membres `name` et `position` sur l’objet [Excel.Worksheet](/javascript/api/excel/excel.worksheet) sont des propriétés scalaires, tandis que `protection` et `tables` sont des propriétés de navigation.</span><span class="sxs-lookup"><span data-stu-id="0152d-163">For example, `name` and `position` members on the [Excel.Worksheet](/javascript/api/excel/excel.worksheet) object are scalar properties, whereas `protection` and `tables` are navigation properties.</span></span>

<span data-ttu-id="0152d-164">Votre complément peut utiliser des propriétés de navigation comme chemin d’accès pour charger des propriétés scalaires spécifiques.</span><span class="sxs-lookup"><span data-stu-id="0152d-164">Your add-in can use navigational properties as a path to load specific scalar properties.</span></span> <span data-ttu-id="0152d-165">Le code suivant met en file d’attente une commande `load` pour le nom de la police utilisée par un objet `Excel.Range`, sans charger d’autres informations.</span><span class="sxs-lookup"><span data-stu-id="0152d-165">The following code queues up a `load` command for the name of the font used by an `Excel.Range` object, without loading any other information.</span></span>

```js
someRange.load("format/font/name")
```

<span data-ttu-id="0152d-166">Vous pouvez également définir les propriétés scalaires d’une propriété de navigation en parcourant le chemin d’accès.</span><span class="sxs-lookup"><span data-stu-id="0152d-166">You can also set the scalar properties of a navigation property by traversing the path.</span></span> <span data-ttu-id="0152d-167">Par exemple, vous pouvez définir la taille de police de `Excel.Range` à l’aide de `someRange.format.font.size = 10;`.</span><span class="sxs-lookup"><span data-stu-id="0152d-167">For example, you could set the font size for an `Excel.Range` by using `someRange.format.font.size = 10;`.</span></span> <span data-ttu-id="0152d-168">Vous n’avez pas besoin de charger la propriété avant de la configurer.</span><span class="sxs-lookup"><span data-stu-id="0152d-168">You don't need to load the property before you set it.</span></span>

<span data-ttu-id="0152d-169">N’oubliez pas que certaines des propriétés sous un objet peuvent avoir le même nom qu’un autre objet.</span><span class="sxs-lookup"><span data-stu-id="0152d-169">Please be aware that some of the properties under an object may have the same name as another object.</span></span> <span data-ttu-id="0152d-170">Par exemple, `format` est une propriété sous l’objet `Excel.Range`, `format` est également un objet.</span><span class="sxs-lookup"><span data-stu-id="0152d-170">For example, `format` is a property under the `Excel.Range` object, but `format` itself is an object as well.</span></span> <span data-ttu-id="0152d-171">Donc, si vous effectuez un appel tel que `range.load("format")`, cela équivaut à `range.format.load()` (une instruction `load()` vide indésirable).</span><span class="sxs-lookup"><span data-stu-id="0152d-171">So, if you make a call such as `range.load("format")`, this is equivalent to `range.format.load()` (an undesirable empty `load()` statement).</span></span> <span data-ttu-id="0152d-172">Pour éviter cela, votre code devrait charger uniquement les nœuds « terminaux » dans une arborescence d’objets.</span><span class="sxs-lookup"><span data-stu-id="0152d-172">To avoid this, your code should only load the "leaf nodes" in an object tree.</span></span>

#### <a name="calling-load-without-parameters-not-recommended"></a><span data-ttu-id="0152d-173">Appel `load` sans paramètres (non recommandé)</span><span class="sxs-lookup"><span data-stu-id="0152d-173">Calling `load` without parameters (not recommended)</span></span>

<span data-ttu-id="0152d-174">Si vous appelez la méthode `load()` sur un objet (ou une collection) sans spécifier de paramètres, toutes les propriétés scalaires de l’objet ou les objets de la collection sont chargées.</span><span class="sxs-lookup"><span data-stu-id="0152d-174">If you call the `load()` method on an object (or collection) without specifying any parameters, all scalar properties of the object or the collection's objects will be loaded.</span></span> <span data-ttu-id="0152d-175">Le chargement des données inutiles ralentit votre complément.</span><span class="sxs-lookup"><span data-stu-id="0152d-175">Loading unneeded data will slow down your add-in.</span></span> <span data-ttu-id="0152d-176">Vous devez toujours spécifier explicitement les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="0152d-176">You should always explicitly specify which properties to load.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="0152d-177">La quantité de données renvoyées par une `load`instruction sans paramètre peut dépasser les limites de taille du service.</span><span class="sxs-lookup"><span data-stu-id="0152d-177">The amount of data returned by a parameter-less `load` statement can exceed the size limits of the service.</span></span> <span data-ttu-id="0152d-178">Pour réduire les risques pesant sur les compléments plus anciens, certaines propriétés ne sont pas renvoyées par `load` sans en faire la demande explicite.</span><span class="sxs-lookup"><span data-stu-id="0152d-178">To reduce the risks to older add-ins, some properties are not returned by `load` without explicitly requesting them.</span></span> <span data-ttu-id="0152d-179">Les propriétés suivantes sont exclues des opérations de chargement suivantes :</span><span class="sxs-lookup"><span data-stu-id="0152d-179">The following properties are excluded from such load operations:</span></span>
>
> * `Excel.Range.numberFormatCategories`

### <a name="clientresult"></a><span data-ttu-id="0152d-180">ClientResult</span><span class="sxs-lookup"><span data-stu-id="0152d-180">ClientResult</span></span>

<span data-ttu-id="0152d-181">Les méthodes utilisées dans les API basées sur la promesse qui renvoient des types possèdent un modèle similaire au modèle `load`/`sync`.</span><span class="sxs-lookup"><span data-stu-id="0152d-181">Methods in the promise-based APIs that return primitive types have a similar pattern to the `load`/`sync` paradigm.</span></span> <span data-ttu-id="0152d-182">Par exemple, `Excel.TableCollection.getCount` obtient le nombre de tableaux dans la collection.</span><span class="sxs-lookup"><span data-stu-id="0152d-182">As an example, `Excel.TableCollection.getCount` gets the number of tables in the collection.</span></span> <span data-ttu-id="0152d-183">`getCount` renvoie un `ClientResult<number>`, ce qui signifie que la propriété `value` dans le [`ClientResult`](/javascript/api/office/officeextension.clientresult) renvoyé est un nombre.</span><span class="sxs-lookup"><span data-stu-id="0152d-183">`getCount` returns a `ClientResult<number>`, meaning the `value` property in the returned [`ClientResult`](/javascript/api/office/officeextension.clientresult) is a number.</span></span> <span data-ttu-id="0152d-184">Votre script ne peut pas accéder à cette valeur tant que `context.sync()` n’est pas appelé.</span><span class="sxs-lookup"><span data-stu-id="0152d-184">Your script can't access that value until `context.sync()` is called.</span></span>

<span data-ttu-id="0152d-185">Le code suivant obtient le nombre total de tableaux dans un feuille de calcul Excel et enregistre ce nombre dans la console.</span><span class="sxs-lookup"><span data-stu-id="0152d-185">The following code gets the total number of tables in an Excel workbook and logs that number to the console.</span></span>

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

### <a name="set"></a><span data-ttu-id="0152d-186">set()</span><span class="sxs-lookup"><span data-stu-id="0152d-186">set()</span></span>

<span data-ttu-id="0152d-187">La définition de propriétés sur un objet avec des propriétés de navigation imbriquées peut être laborieuse.</span><span class="sxs-lookup"><span data-stu-id="0152d-187">Setting properties on an object with nested navigation properties can be cumbersome.</span></span> <span data-ttu-id="0152d-188">Au lieu de définir des propriétés individuelles à l’aide de chemins de navigation comme décrit ci-dessus, vous pouvez utiliser la méthode `object.set()` disponible sur les objets dans les API JavaScript basées sur une promesse.</span><span class="sxs-lookup"><span data-stu-id="0152d-188">As an alternative to setting individual properties using navigation paths as described above, you can use the `object.set()` method that is available on objects in the promise-based JavaScript APIs.</span></span> <span data-ttu-id="0152d-189">Grâce à cette méthode, vous pouvez définir plusieurs propriétés d’un objet à la fois en transmettant soit un autre objet du même type Office.js, soit un objet JavaScript avec des propriétés structurées comme celles de l’objet sur lequel la méthode est appelée.</span><span class="sxs-lookup"><span data-stu-id="0152d-189">With this method, you can set multiple properties of an object at once by passing either another object of the same Office.js type or a JavaScript object with properties that are structured like the properties of the object on which the method is called.</span></span>

<span data-ttu-id="0152d-p124">L’exemple de code suivant définit plusieurs propriétés de mise en forme d’une plage en appelant la méthode `set()` et en transmettant un objet JavaScript avec des noms et des types de propriétés reflétant la structure des propriétés dans l’objet `Range`. Cet exemple part du principe que des données sont présentes dans la plage **B2:E2**.</span><span class="sxs-lookup"><span data-stu-id="0152d-p124">The following code sample sets several format properties of a range by calling the `set()` method and passing in a JavaScript object with property names and types that mirror the structure of properties in the `Range` object. This example assumes that there is data in range **B2:E2**.</span></span>

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

### <a name="some-properties-cannot-be-set-directly"></a><span data-ttu-id="0152d-192">Certaines propriétés ne peuvent pas être définies directement</span><span class="sxs-lookup"><span data-stu-id="0152d-192">Some properties cannot be set directly</span></span>

<span data-ttu-id="0152d-193">Certaines propriétés ne peuvent pas être définies, même si elles sont accessibles en écriture.</span><span class="sxs-lookup"><span data-stu-id="0152d-193">Some properties cannot be set, despite being writable.</span></span> <span data-ttu-id="0152d-194">Ces propriétés font partie d’une propriété parente qui doit être définie en tant qu’objet unique.</span><span class="sxs-lookup"><span data-stu-id="0152d-194">These properties are part of a parent property that must be set as a single object.</span></span> <span data-ttu-id="0152d-195">En effet, cette propriété parente s’appuie sur les sous-propriétés ayant des relations logiques spécifiques.</span><span class="sxs-lookup"><span data-stu-id="0152d-195">This is because that parent property relies on the subproperties having specific, logical relationships.</span></span> <span data-ttu-id="0152d-196">Ces propriétés parentes doivent être définies à l’aide de la notation littérale de l’objet pour définir l’intégralité de l’objet, plutôt que de définir les sous-propriétés individuelles de cet objet.</span><span class="sxs-lookup"><span data-stu-id="0152d-196">These parent properties must be set using object literal notation to set the entire object, instead of setting that object's individual subproperties.</span></span> <span data-ttu-id="0152d-197">Un exemple de ce modèle est trouvé dans [PageLayout](/javascript/api/excel/excel.pagelayout).</span><span class="sxs-lookup"><span data-stu-id="0152d-197">One example of this is found in [PageLayout](/javascript/api/excel/excel.pagelayout).</span></span> <span data-ttu-id="0152d-198">La propriété `zoom` doit être définie avec un objet [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) unique, comme illustré ici :</span><span class="sxs-lookup"><span data-stu-id="0152d-198">The `zoom` property must be set with a single [PageLayoutZoomOptions](/javascript/api/excel/excel.pagelayoutzoomoptions) object, as shown here:</span></span>

```js
// PageLayout.zoom.scale must be set by assigning PageLayout.zoom to a PageLayoutZoomOptions object.
sheet.pageLayout.zoom = { scale: 200 };
```

<span data-ttu-id="0152d-199">Dans l’exemple précédent, vous ***ne pouvez pas*** affecter directement une valeur à `zoom` : `sheet.pageLayout.zoom.scale = 200;`.</span><span class="sxs-lookup"><span data-stu-id="0152d-199">In the previous example, you would ***not*** be able to directly assign `zoom` a value: `sheet.pageLayout.zoom.scale = 200;`.</span></span> <span data-ttu-id="0152d-200">Cette instruction génère une erreur, car `zoom` n’est pas chargé.</span><span class="sxs-lookup"><span data-stu-id="0152d-200">That statement throws an error because `zoom` is not loaded.</span></span> <span data-ttu-id="0152d-201">Même si `zoom` était chargé, l’ensemble d’échelles n’est pas pris en compte.</span><span class="sxs-lookup"><span data-stu-id="0152d-201">Even if `zoom` were to be loaded, the set of scale will not take effect.</span></span> <span data-ttu-id="0152d-202">Toutes les opérations de contexte se produisent sur `zoom`, elles actualisent l’objet proxy du complément et remplacement des valeurs définies localement.</span><span class="sxs-lookup"><span data-stu-id="0152d-202">All context operations happen on `zoom`, refreshing the proxy object in the add-in and overwriting locally set values.</span></span>

<span data-ttu-id="0152d-203">Ce comportement diffère des [propriétés de navigation](application-specific-api-model.md#scalar-and-navigation-properties) telles que [Range.format](/javascript/api/excel/excel.range#format).</span><span class="sxs-lookup"><span data-stu-id="0152d-203">This behavior differs from [navigational properties](application-specific-api-model.md#scalar-and-navigation-properties) like [Range.format](/javascript/api/excel/excel.range#format).</span></span> <span data-ttu-id="0152d-204">Les propriétés de `format` peuvent être définies à l’aide de la navigation d’objets, comme illustré ici :</span><span class="sxs-lookup"><span data-stu-id="0152d-204">Properties of `format` can be set using object navigation, as shown here:</span></span>

```js
// This will set the font size on the range during the next `content.sync()`.
range.format.font.size = 10;
```

<span data-ttu-id="0152d-205">Vous pouvez identifier une propriété qui ne peut pas avoir ses sous-propriétés définies directement en consultant son modificateur en lecture seule.</span><span class="sxs-lookup"><span data-stu-id="0152d-205">You can identify a property that cannot have its subproperties directly set by checking its read-only modifier.</span></span> <span data-ttu-id="0152d-206">Toutes les propriétés en lecture seule peuvent avoir leurs sous-propriétés sans lecture seule directement définit.</span><span class="sxs-lookup"><span data-stu-id="0152d-206">All read-only properties can have their non-read-only subproperties directly set.</span></span> <span data-ttu-id="0152d-207">Les propriétés disponibles en écriture comme `PageLayout.zoom`, par exemple, doivent être définies avec un objet de ce niveau.</span><span class="sxs-lookup"><span data-stu-id="0152d-207">Writeable properties like `PageLayout.zoom` must be set with an object at that level.</span></span> <span data-ttu-id="0152d-208">En Résumé :</span><span class="sxs-lookup"><span data-stu-id="0152d-208">In summary:</span></span>

- <span data-ttu-id="0152d-209">Propriété en lecture seule : les sous-propriétés peuvent être définies via la navigation.</span><span class="sxs-lookup"><span data-stu-id="0152d-209">Read-only property: Subproperties can be set through navigation.</span></span>
- <span data-ttu-id="0152d-210">Propriété accessibles en écriture : les sous-propriétés ne peuvent pas être définies via la navigation (elles doivent être définies dans le cadre de l’affectation d’objet parent initiale).</span><span class="sxs-lookup"><span data-stu-id="0152d-210">Writable property: Subproperties cannot be set through navigation (must be set as part of the initial parent object assignment).</span></span>



## <a name="42ornullobject-methods-and-properties"></a><span data-ttu-id="0152d-211">Méthodes et propriétés de &#42;OrNullObject</span><span class="sxs-lookup"><span data-stu-id="0152d-211">&#42;OrNullObject methods and properties</span></span>

<span data-ttu-id="0152d-212">Certaines méthodes et propriétés d’accessoires ajoutent une exception lorsque l’objet souhaité n’existe pas.</span><span class="sxs-lookup"><span data-stu-id="0152d-212">Some accessor methods and properties throw an exception when the desired object doesn't exist.</span></span> <span data-ttu-id="0152d-213">Par exemple, si vous tentez d’obtenir une feuille de calcul Excel en spécifiant le nom d’une feuille de calcul qui n’existe pas dans le classeur, la méthode `getItem()` renvoie une exception `ItemNotFound`.</span><span class="sxs-lookup"><span data-stu-id="0152d-213">For example, if you attempt to get an Excel worksheet by specifying a worksheet name that isn't in the workbook, the `getItem()` method throws an `ItemNotFound` exception.</span></span> <span data-ttu-id="0152d-214">Les bibliothèques spécifiques de l’application permettent à votre code de tester l’existence d’entités de document sans exiger de code de gestion d’exceptions.</span><span class="sxs-lookup"><span data-stu-id="0152d-214">The application-specific libraries provide a way for your code to test for the existence of document entities without requiring exception handling code.</span></span> <span data-ttu-id="0152d-215">Cela est possible à l’aide des variantes `*OrNullObject` de méthodes et de propriétés.</span><span class="sxs-lookup"><span data-stu-id="0152d-215">This is accomplished by using the `*OrNullObject` variations of methods and properties.</span></span> <span data-ttu-id="0152d-216">Ces variantes renvoient un objet dont la propriété `isNullObject` est définie sur `true`, si l’élément spécifié n’existe pas, plutôt que de renvoyer une exception.</span><span class="sxs-lookup"><span data-stu-id="0152d-216">These variations return an object whose `isNullObject` property is set to `true`, if the specified item doesn't exist, rather than throwing an exception.</span></span>

<span data-ttu-id="0152d-217">Par exemple, vous pouvez appeler la méthode `getItemOrNullObject()` sur une collection telle que **Feuilles de calcul** pour récupérer un élément de la collection.</span><span class="sxs-lookup"><span data-stu-id="0152d-217">For example, you can call the `getItemOrNullObject()` method on a collection such as **Worksheets** to retrieve an item from the collection.</span></span> <span data-ttu-id="0152d-218">La méthode `getItemOrNullObject()` renvoie l’élément spécifié s’il existe. sinon, il renvoie un objet dont la propriété `isNullObject` est définie sur `true`.</span><span class="sxs-lookup"><span data-stu-id="0152d-218">The `getItemOrNullObject()` method returns the specified item if it exists; otherwise, it returns an object whose `isNullObject` property is set to `true`.</span></span> <span data-ttu-id="0152d-219">Votre code peut ensuite évaluer cette propriété pour déterminer si l’objet existe.</span><span class="sxs-lookup"><span data-stu-id="0152d-219">Your code can then evaluate this property to determine whether the object exists.</span></span>

> [!NOTE]
> <span data-ttu-id="0152d-220">Les variantes `*OrNullObject` ne renvoient jamais la valeur JavaScript `null`.</span><span class="sxs-lookup"><span data-stu-id="0152d-220">The `*OrNullObject` variations do not ever return the JavaScript value `null`.</span></span> <span data-ttu-id="0152d-221">Ils renvoient des objets proxy Office ordinaires.</span><span class="sxs-lookup"><span data-stu-id="0152d-221">They return ordinary Office proxy objects.</span></span> <span data-ttu-id="0152d-222">Si l’entité que l’objet représente n’existe pas, la propriété `isNullObject` de l’objet est définie sur `true`.</span><span class="sxs-lookup"><span data-stu-id="0152d-222">If the the entity that the object represents does not exist then the `isNullObject` property of the object is set to `true`.</span></span> <span data-ttu-id="0152d-223">Ne testez pas l’objet renvoyé pour nullité ou fausseté.</span><span class="sxs-lookup"><span data-stu-id="0152d-223">Do not test the returned object for nullity or falsity.</span></span> <span data-ttu-id="0152d-224">Ce n’est jamais `null`, `false` ou `undefined`.</span><span class="sxs-lookup"><span data-stu-id="0152d-224">It is never `null`, `false`, or `undefined`.</span></span>

<span data-ttu-id="0152d-225">L’exemple de code suivant tente de récupérer une feuille de calcul Excel nommée « Données » à l’aide de la méthode `getItemOrNullObject()`.</span><span class="sxs-lookup"><span data-stu-id="0152d-225">The following code sample attempts to retrieve an Excel worksheet named "Data" by using the `getItemOrNullObject()` method.</span></span> <span data-ttu-id="0152d-226">Si une feuille de calcul avec ce nom n’existe pas, une nouvelle feuille est créée.</span><span class="sxs-lookup"><span data-stu-id="0152d-226">If a worksheet with that name does not exist, a new sheet is created.</span></span> <span data-ttu-id="0152d-227">Notez que le code ne charge pas la propriété `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="0152d-227">Note that the code does not load the `isNullObject` property.</span></span> <span data-ttu-id="0152d-228">Office charge automatiquement cette propriété lorsque `context.sync` est appelé. Vous n’avez donc pas besoin de la charger explicitement avec quelque chose comme `datasheet.load('isNullObject')`.</span><span class="sxs-lookup"><span data-stu-id="0152d-228">Office automatically loads this property when `context.sync` is called, so you do not need to explicitly load it with something like `datasheet.load('isNullObject')`.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="0152d-229">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="0152d-229">See also</span></span>

* [<span data-ttu-id="0152d-230">Modèle d’objet API JavaScript courant</span><span class="sxs-lookup"><span data-stu-id="0152d-230">Common JavaScript API object model</span></span>](office-javascript-api-object-model.md)
* [<span data-ttu-id="0152d-231">Limites des ressources et optimisation des performances pour les compléments Office</span><span class="sxs-lookup"><span data-stu-id="0152d-231">Resource limits and performance optimization for Office Add-ins</span></span>](../concepts/resource-limits-and-performance-optimization.md)
