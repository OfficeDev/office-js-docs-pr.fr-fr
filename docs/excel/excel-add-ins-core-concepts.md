---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Excel
description: Utilisez l’API JavaScript pour Excel afin de créer des compléments pour Excel.
ms.date: 11/29/2018
localization_priority: Priority
ms.openlocfilehash: b04ed4053f0cc16c38d69cf62942559e1b4f514a
ms.sourcegitcommit: d1aa7201820176ed986b9f00bb9c88e055906c77
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 01/23/2019
ms.locfileid: "29388702"
---
# <a name="fundamental-programming-concepts-with-the-excel-javascript-api"></a><span data-ttu-id="f85c3-103">Concepts fondamentaux de programmation avec l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-103">Fundamental programming concepts with the Excel JavaScript API</span></span>
 
<span data-ttu-id="f85c3-104">Cet article décrit comment utiliser l’[API JavaScript pour Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) afin de créer des compléments pour Excel 2016 ou versions ultérieures.</span><span class="sxs-lookup"><span data-stu-id="f85c3-104">This article describes how to use the [Excel JavaScript API](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview) to build add-ins for Excel 2016 or later.</span></span> <span data-ttu-id="f85c3-105">Il présente les concepts fondamentaux de l’utilisation des API et fournit des conseils pour effectuer des tâches spécifiques, comme la lecture ou l’écriture d’une grande plage, la mise à jour de toutes les cellules d’une plage, et bien plus encore.</span><span class="sxs-lookup"><span data-stu-id="f85c3-105">It introduces core concepts that are fundamental to using the API and provides guidance for performing specific tasks such as reading or writing to a large range, updating all cells in range, and more.</span></span>

## <a name="asynchronous-nature-of-excel-apis"></a><span data-ttu-id="f85c3-106">Nature asynchrone des API Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-106">Asynchronous nature of Excel APIs</span></span>

<span data-ttu-id="f85c3-107">Les compléments Excel web s’exécutent dans un conteneur de navigateurs qui est incorporé dans l’application Office sur les plateformes basées sur un bureau, comme Office pour Windows, et s’exécute à l’intérieur d’un fichier iFrame HTML dans Office Online.</span><span class="sxs-lookup"><span data-stu-id="f85c3-107">The web-based Excel add-ins run inside a browser container that is embedded within the Office application on desktop-based platforms such as Office for Windows and runs inside an HTML iFrame in Office Online.</span></span> <span data-ttu-id="f85c3-108">En raison de problèmes de performances, il n’est pas possible d’activer l’API Office.js afin d’interagir de manière synchrone avec l’hôte Excel sur toutes les plateformes prises en charge.</span><span class="sxs-lookup"><span data-stu-id="f85c3-108">Enabling the Office.js API to interact synchronously with the Excel host across all supported platforms is not feasible due to performance considerations.</span></span> <span data-ttu-id="f85c3-109">Par conséquent, l’appel de l’API **sync()** dans Office.js renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) qui est résolue lorsque l’application Excel termine les actions de lecture ou d’écriture demandées.</span><span class="sxs-lookup"><span data-stu-id="f85c3-109">Therefore, the **sync()** API call in Office.js returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) that is resolved when the Excel application completes the requested read or write actions.</span></span> <span data-ttu-id="f85c3-110">En outre, vous pouvez mettre en file d’attente plusieurs actions, comme la définition des propriétés ou l’appel de méthodes, et les exécuter en tant que lot de commandes avec un seul appel à **sync()**, au lieu d’envoyer une demande distincte pour chaque action.</span><span class="sxs-lookup"><span data-stu-id="f85c3-110">Also, you can queue up multiple actions, such as setting properties or invoking methods, and run them as a batch of commands with a single call to **sync()**, rather than sending a separate request for each action.</span></span> <span data-ttu-id="f85c3-111">Les sections suivantes décrivent la façon d’y parvenir à l’aide des API **Excel.run()** et **sync()**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-111">The following sections describe how to accomplish this using the **Excel.run()** and **sync()** APIs.</span></span>
 
## <a name="excelrun"></a><span data-ttu-id="f85c3-112">Excel.run</span><span class="sxs-lookup"><span data-stu-id="f85c3-112">Excel.run</span></span>
 
<span data-ttu-id="f85c3-113">**Excel.Run** exécute une fonction dans laquelle vous spécifiez les actions à effectuer concernant le modèle objet Excel.</span><span class="sxs-lookup"><span data-stu-id="f85c3-113">**Excel.run** executes a function where you specify the actions to perform against the Excel object model.</span></span> <span data-ttu-id="f85c3-114">**Excel.Run** crée automatiquement un contexte de la demande que vous pouvez utiliser pour interagir avec des objets Excel.</span><span class="sxs-lookup"><span data-stu-id="f85c3-114">**Excel.run** automatically creates a request context that you can use to interact with Excel objects.</span></span> <span data-ttu-id="f85c3-115">Lorsque l’API **Excel.run** a fini, une promesse est résolue, et tous les objets alloués lors de l’exécution sont automatiquement publiés.</span><span class="sxs-lookup"><span data-stu-id="f85c3-115">When **Excel.run** completes, a promise is resolved, and any objects that were allocated at runtime are automatically released.</span></span>
 
<span data-ttu-id="f85c3-116">L’exemple suivant montre comment utiliser **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-116">The following example shows how to use **Excel.run**.</span></span> <span data-ttu-id="f85c3-117">L’instruction catch capture et enregistre les erreurs qui se produisent au sein de **Excel.run**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-117">The catch statement catches and logs errors that occur within the **Excel.run**.</span></span>
 
```js
Excel.run(function (context) {
  // You can use the Excel JavaScript API here in the batch function
  // to execute actions on the Excel object model.
  console.log('Your code goes here.');
}).catch(function (error) {
  console.log('error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```

### <a name="run-options"></a><span data-ttu-id="f85c3-118">Options d’exécution</span><span class="sxs-lookup"><span data-stu-id="f85c3-118">Run options</span></span>

<span data-ttu-id="f85c3-119">**Excel.Run** est associé à une surcharge liée à un objet [RunOptions](/javascript/api/excel/excel.runoptions).</span><span class="sxs-lookup"><span data-stu-id="f85c3-119">**Excel.run** has an overload that takes in a [RunOptions](/javascript/api/excel/excel.runoptions) object.</span></span> <span data-ttu-id="f85c3-120">Celui-ci contient un ensemble de propriétés qui ont une incidence sur le comportement de la plateforme lorsque la fonction est en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="f85c3-120">This contains a set of properties that affect platform behavior when the function runs.</span></span> <span data-ttu-id="f85c3-121">La propriété suivante est actuellement prise en charge :</span><span class="sxs-lookup"><span data-stu-id="f85c3-121">The following property is currently supported:</span></span>

 - <span data-ttu-id="f85c3-122">`delayForCellEdit` : détermine si Excel diffère la demande de lot jusqu'à ce que l’utilisateur quitte le mode de modification de cellule.</span><span class="sxs-lookup"><span data-stu-id="f85c3-122">`delayForCellEdit`: Determines whether Excel delays the batch request until the user exits cell edit mode.</span></span> <span data-ttu-id="f85c3-123">Lorsque la valeur est **true**, la demande de lot est différée et s’exécute lorsque l’utilisateur quitte le mode de modification de cellule.</span><span class="sxs-lookup"><span data-stu-id="f85c3-123">When **true**, the batch request is delayed and runs when the user exits cell edit mode.</span></span> <span data-ttu-id="f85c3-124">Lorsque la valeur est **false**, la demande de lot échoue automatiquement si l’utilisateur est en mode de modification de cellule (entraînant une erreur de contact de l’utilisateur).</span><span class="sxs-lookup"><span data-stu-id="f85c3-124">When **false**, the batch request automatically fails if the user is in cell edit mode (causing an error to reach the user).</span></span> <span data-ttu-id="f85c3-125">Le comportement par défaut sans propriété `delayForCellEdit` spécifiée est identique au comportement lorsque la valeur est **false**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-125">The default behavior with no `delayForCellEdit` property specified is equivalent to when it is **false**.</span></span>

```js
Excel.run({ delayForCellEdit: true }, function (context) { ... })
```

## <a name="request-context"></a><span data-ttu-id="f85c3-126">Contexte de demande</span><span class="sxs-lookup"><span data-stu-id="f85c3-126">Request context</span></span>
 
<span data-ttu-id="f85c3-p107">Excel et votre complément sont exécutés dans deux processus distincts. Dans la mesure où ils utilisent des environnements d’exécution différents, les compléments Excel nécessitent un objet **RequestContext** afin de connecter votre complément aux objets dans Excel, tels que les feuilles de calcul, les plages, les graphiques et les tableaux.</span><span class="sxs-lookup"><span data-stu-id="f85c3-p107">Excel and your add-in run in two different processes. Since they use different runtime environments, Excel add-ins require a **RequestContext** object in order to connect your add-in to objects in Excel such as worksheets, ranges, charts, and tables.</span></span>
 
## <a name="proxy-objects"></a><span data-ttu-id="f85c3-129">Objets de proxy</span><span class="sxs-lookup"><span data-stu-id="f85c3-129">Proxy objects</span></span>
 
<span data-ttu-id="f85c3-130">Les objets JavaScript pour Excel que vous déclarez et utilisez dans un complément sont des objets proxy.</span><span class="sxs-lookup"><span data-stu-id="f85c3-130">The Excel JavaScript objects that you declare and use in an add-in are proxy objects.</span></span> <span data-ttu-id="f85c3-131">Les méthodes que vous appelez ou les propriétés que vous définissez ou chargez sur les objets proxy sont simplement ajoutées à une file d’attente de commandes en attente.</span><span class="sxs-lookup"><span data-stu-id="f85c3-131">Any methods that you invoke or properties that you set or load on proxy objects are simply added to a queue of pending commands.</span></span> <span data-ttu-id="f85c3-132">Lorsque vous appelez la méthode **sync()** sur le contexte de demande (par exemple, `context.sync()`), les commandes en attente sont envoyées vers Excel et sont exécutées.</span><span class="sxs-lookup"><span data-stu-id="f85c3-132">When you call the **sync()** method on the request context (for example, `context.sync()`), the queued commands are dispatched to Excel and run.</span></span> <span data-ttu-id="f85c3-133">L’API JavaScript pour Excel est fondamentalement centrée sur les lots.</span><span class="sxs-lookup"><span data-stu-id="f85c3-133">The Excel JavaScript API is fundamentally batch-centric.</span></span> <span data-ttu-id="f85c3-134">Vous pouvez mettre en file d’attente autant de modifications que vous le souhaitez dans le contexte de la demande, puis appeler la méthode **sync()** pour exécuter le lot de commandes mises en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="f85c3-134">You can queue up as many changes as you wish on the request context, and then call the **sync()** method to run the batch of queued commands.</span></span>
 
<span data-ttu-id="f85c3-135">Par exemple, l’extrait de code suivant déclare l’objet JavaScript local **selectedRange** pour référencer une plage sélectionnée dans le document Excel, puis définit des propriétés sur cet objet.</span><span class="sxs-lookup"><span data-stu-id="f85c3-135">For example, the following code snippet declares the local JavaScript object **selectedRange** to reference a selected range in the Excel document, and then sets some properties on that object.</span></span> <span data-ttu-id="f85c3-136">L’objet **selectedRange** est un objet proxy. Les propriétés définies et la méthode appelée sur cet objet ne seront pas répercutées dans le document Excel tant que votre complément n’a pas appelé **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-136">The **selectedRange** object is a proxy object, so the properties that are set and method that is invoked on that object will not be reflected in the Excel document until your add-in calls **context.sync()**.</span></span>
 
```js
const selectedRange = context.workbook.getSelectedRange();
selectedRange.format.fill.color = "#4472C4";
selectedRange.format.font.color = "white";
selectedRange.format.autofitColumns();
```
 
### <a name="sync"></a><span data-ttu-id="f85c3-137">Sync</span><span class="sxs-lookup"><span data-stu-id="f85c3-137">sync()</span></span>
 
<span data-ttu-id="f85c3-138">Tout appel de la méthode **sync()** concernant le contexte de demande synchronise l’état entre les objets proxy et les objets du document Excel.</span><span class="sxs-lookup"><span data-stu-id="f85c3-138">Calling the **sync()** method on the request context synchronizes the state between proxy objects and objects in the Excel document.</span></span> <span data-ttu-id="f85c3-139">La méthode **sync()** exécute les commandes mises en file d’attente concernant le contexte de demande et récupère des valeurs pour les propriétés qui doivent être chargées dans les objets proxy.</span><span class="sxs-lookup"><span data-stu-id="f85c3-139">The **sync()** method runs any commands that are queued on the request context and retrieves values for any properties that should be loaded on the proxy objects.</span></span> <span data-ttu-id="f85c3-140">La méthode **sync()** est exécutée de façon asynchrone et renvoie une [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), qui est résolue lorsque la méthode **sync()** est terminée.</span><span class="sxs-lookup"><span data-stu-id="f85c3-140">The **sync()** method executes asynchronously and returns a [promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise), which is resolved when the **sync()** method completes.</span></span>
 
<span data-ttu-id="f85c3-141">L’exemple suivant montre une fonction de traitement par lot qui définit un objet proxy JavaScript local (**selectedRange**), charge une propriété de cet objet et utilise ensuite le modèle de promesses JavaScript pour appeler **context.sync()** afin de synchroniser l’état entre les objets proxy et les objets du document Excel.</span><span class="sxs-lookup"><span data-stu-id="f85c3-141">The following example shows a batch function that defines a local JavaScript proxy object (**selectedRange**), loads a property of that object, and then uses the JavaScript Promises pattern to call **context.sync()** to synchronize the state between proxy objects and objects in the Excel document.</span></span>
 
```js
Excel.run(function (context) {
  const selectedRange = context.workbook.getSelectedRange();
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
 
<span data-ttu-id="f85c3-142">Dans l’exemple précédent, l’objet **selectedRange** est défini et sa propriété **address** est chargée quand l’élément **context.sync()** est appelé.</span><span class="sxs-lookup"><span data-stu-id="f85c3-142">In the previous example, **selectedRange** is set and its **address** property is loaded when **context.sync()** is called.</span></span>
 
<span data-ttu-id="f85c3-143">Étant donné que **sync()** est une opération asynchrone qui renvoie une promesse, vous devez toujours **renvoyer** la promesse (dans JavaScript).</span><span class="sxs-lookup"><span data-stu-id="f85c3-143">Because **sync()** is an asynchronous operation that returns a promise, you should always **return** the promise (in JavaScript).</span></span> <span data-ttu-id="f85c3-144">Cela garantit que l’opération **sync()** se termine avant que le script continue à s’exécuter.</span><span class="sxs-lookup"><span data-stu-id="f85c3-144">Doing so ensures that the **sync()** operation completes before the script continues to run.</span></span> <span data-ttu-id="f85c3-145">Pour plus d’informations sur l’optimisation des performances avec **sync()**, consultez la rubrique [Optimisation des performances à l’aide de l’API JavaScript d’Excel](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span><span class="sxs-lookup"><span data-stu-id="f85c3-145">For more information about optimizing performance with **sync()**, see [Excel JavaScript API performance optimization](https://docs.microsoft.com/office/dev/add-ins/excel/performance).</span></span>
 
### <a name="load"></a><span data-ttu-id="f85c3-146">load()</span><span class="sxs-lookup"><span data-stu-id="f85c3-146">load()</span></span>
 
<span data-ttu-id="f85c3-147">Avant que vous puissiez lire les propriétés d’un objet proxy, vous devez charger explicitement les propriétés pour remplir l’objet proxy avec des données à partir du document Excel, puis appeler **context.sync()**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-147">Before you can read the properties of a proxy object, you must explicitly load the properties to populate the proxy object with data from the Excel document, and then call **context.sync()**.</span></span> <span data-ttu-id="f85c3-148">Par exemple, si vous créez un objet proxy pour référencer une plage sélectionnée, puis que vous voulez lire la propriété **address** de la plage sélectionnée, vous devez charger la propriété **address** avant de pouvoir la lire.</span><span class="sxs-lookup"><span data-stu-id="f85c3-148">For example, if you create a proxy object to reference a selected range, and then want to read the selected range's **address** property, you need to load the **address** property before you can read it.</span></span> <span data-ttu-id="f85c3-149">Pour demander le chargement de propriétés d’un objet, appelez la méthode **load()** sur l’objet et spécifiez les propriétés à charger.</span><span class="sxs-lookup"><span data-stu-id="f85c3-149">To request properties of a proxy object be loaded, call the **load()** method on the object and specify the properties to load.</span></span> 

> [!NOTE]
> <span data-ttu-id="f85c3-150">Si vous appelez uniquement des méthodes ou définissez des propriétés sur un objet proxy, il est inutile d’appeler la méthode **load()**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-150">If you are only calling methods or setting properties on a proxy object, you do not need to call the **load()** method.</span></span> <span data-ttu-id="f85c3-151">La méthode **load()** n’est nécessaire que lorsque vous souhaitez lire les propriétés sur un objet proxy.</span><span class="sxs-lookup"><span data-stu-id="f85c3-151">The **load()** method is only required when you want to read properties on a proxy object.</span></span>
 
<span data-ttu-id="f85c3-p114">À l’instar des demandes de définition de propriétés ou d’appel de méthodes sur des objets proxy, des demandes de chargement de propriétés sur des objets proxy sont ajoutées à la file d’attente des commandes sur le contexte de demande, qui s’exécutera la prochaine fois que vous appellerez la méthode **sync()**. Vous pouvez mettre en file d’attente autant d’appels **load()** sur le contexte de la demande que nécessaire.</span><span class="sxs-lookup"><span data-stu-id="f85c3-p114">Just like requests to set properties or invoke methods on proxy objects, requests to load properties on proxy objects get added to the queue of pending commands on the request context, which will run the next time you call the **sync()** method. You can queue up as many **load()** calls on the request context as necessary.</span></span>
 
<span data-ttu-id="f85c3-154">Dans l’exemple suivant, seules les propriétés spécifiques de la plage sont chargées.</span><span class="sxs-lookup"><span data-stu-id="f85c3-154">In the following example, only specific properties of the range are loaded.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:B2';
  const myRange = context.workbook.worksheets.getItem(sheetName).getRange(rangeAddress);
 
  myRange.load(['address', 'format/*', 'format/fill', 'entireRow' ]);
 
  return context.sync()
    .then(function () {
      console.log (myRange.address);              // ok
      console.log (myRange.format.wrapText);      // ok
      console.log (myRange.format.fill.color);    // ok
      //console.log (myRange.format.font.color);  // not ok as it was not loaded
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
 
<span data-ttu-id="f85c3-155">Comme `format/font` n’est pas spécifié dans l’appel à **myRange.load()**, la propriété `format.font.color` ne peut pas être lue dans l’exemple précédent.</span><span class="sxs-lookup"><span data-stu-id="f85c3-155">In the previous example, because `format/font` is not specified in the call to **myRange.load()**, the `format.font.color` property cannot be read.</span></span>

<span data-ttu-id="f85c3-156">Pour optimiser les performances, vous devez spécifier explicitement les propriétés et les relations à charger lorsque vous utilisez la méthode **load()** sur un objet, comme abordé dans la rubrique [Optimisation des performances à l’aide de l’API JavaScript d’Excel](performance.md).</span><span class="sxs-lookup"><span data-stu-id="f85c3-156">To optimize performance, you should explicitly specify the properties and relationships to load when using the **load()** method on an object, as covered in [Excel JavaScript API performance optimizations](performance.md).</span></span> <span data-ttu-id="f85c3-157">Pour plus d’informations sur la méthode **load()**, consultez la rubrique [Concepts avancés de programmation avec l’API JavaScript Excel](excel-add-ins-advanced-concepts.md).</span><span class="sxs-lookup"><span data-stu-id="f85c3-157">For more information about the **load()** method, see [Advanced programming concepts with the Excel JavaScript API](excel-add-ins-advanced-concepts.md).</span></span>

## <a name="null-or-blank-property-values"></a><span data-ttu-id="f85c3-158">valeurs de propriété null ou vides</span><span class="sxs-lookup"><span data-stu-id="f85c3-158">null or blank property values</span></span>
 
### <a name="null-input-in-2-d-array"></a><span data-ttu-id="f85c3-159">entrée de valeurs null dans un tableau 2D</span><span class="sxs-lookup"><span data-stu-id="f85c3-159">null input in 2-D Array</span></span>
 
<span data-ttu-id="f85c3-160">Dans Excel, une plage est représentée par un tableau 2D, où les lignes représentent la première dimension et les colonnes la deuxième.</span><span class="sxs-lookup"><span data-stu-id="f85c3-160">In Excel, a range is represented by a 2-D array, where the first dimension is rows and the second dimension is columns.</span></span> <span data-ttu-id="f85c3-161">Pour définir des valeurs, un format de nombre ou une formule uniquement pour des cellules spécifiques dans une plage, spécifiez des valeurs, un format de nombre ou une formule pour ces cellules dans le tableau 2D, et indiquez `null` pour toutes les autres cellules du tableau 2D.</span><span class="sxs-lookup"><span data-stu-id="f85c3-161">To set values, number format, or formula for only specific cells within a range, specify the values, number format, or formula for those cells in the 2-D array, and specify `null` for all other cells in the 2-D array.</span></span>
 
<span data-ttu-id="f85c3-162">Par exemple, pour mettre à jour le format de nombre pour une seule cellule dans une plage et conserver le format de nombre existant pour toutes les autres cellules de la plage, spécifiez le nouveau format de nombre de la cellule à mettre à jour, puis spécifiez `null` pour toutes les autres cellules.</span><span class="sxs-lookup"><span data-stu-id="f85c3-162">For example, to update the number format for only one cell within a range, and retain the existing number format for all other cells in the range, specify the new number format for the cell to update, and specify `null` for all other cells.</span></span> <span data-ttu-id="f85c3-163">L’extrait de code suivant définit un nouveau format de nombre pour la quatrième cellule de la plage et ne modifie pas le format de nombre pour les trois premières cellules de la plage.</span><span class="sxs-lookup"><span data-stu-id="f85c3-163">The following code snippet sets a new number format for the fourth cell in the range, and leaves the number format unchanged for the first three cells in the range.</span></span>
 
```js
range.values = [['Eurasia', '29.96', '0.25', '15-Feb' ]];
range.numberFormat = [[null, null, null, 'm/d/yyyy;@']];
```
 
### <a name="null-input-for-a-property"></a><span data-ttu-id="f85c3-164">Entrée null pour une propriété</span><span class="sxs-lookup"><span data-stu-id="f85c3-164">null input for a property</span></span>
 
<span data-ttu-id="f85c3-p118">`null` n’est pas une entrée valide pour une propriété unique. Par exemple, l’extrait de code suivant n’est pas valide, car la propriété **values** de la plage ne peut pas être définie sur `null`.</span><span class="sxs-lookup"><span data-stu-id="f85c3-p118">`null` is not a valid input for single property. For example, the following code snippet is not valid, as the **values** property of the range cannot be set to `null`.</span></span>
 
```js
range.values = null;
```
 
<span data-ttu-id="f85c3-167">De même, l’extrait de code suivant n’est pas valide, car `null` n’est pas une valeur valide pour la propriété **color**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-167">Likewise, the following code snippet is not valid, as `null` is not a valid value for the **color** property.</span></span>
 
```js
range.format.fill.color =  null;
```
 
### <a name="null-property-values-in-the-response"></a><span data-ttu-id="f85c3-168">valeurs de la propriété Null dans la réponse</span><span class="sxs-lookup"><span data-stu-id="f85c3-168">null property values in the response</span></span>
 
<span data-ttu-id="f85c3-169">Les propriétés de mise en forme comme `size` et `color` contiendront des valeurs `null` dans la réponse lorsque différentes valeurs existent dans la plage spécifiée.</span><span class="sxs-lookup"><span data-stu-id="f85c3-169">Formatting properties such as `size` and `color` will contain `null` values in the response when different values exist in the specified range.</span></span> <span data-ttu-id="f85c3-170">Par exemple, si vous récupérez une plage et chargez sa propriété `format.font.color` :</span><span class="sxs-lookup"><span data-stu-id="f85c3-170">For example, if you retrieve a range and load its `format.font.color` property:</span></span>
 
* <span data-ttu-id="f85c3-171">Si toutes les cellules de la plage ont la même couleur de police, `range.format.font.color` spécifie cette couleur.</span><span class="sxs-lookup"><span data-stu-id="f85c3-171">If all cells in the range have the same font color, `range.format.font.color` specifies that color.</span></span>
* <span data-ttu-id="f85c3-172">Si plusieurs couleurs de police sont présentes dans la plage, `range.format.font.color` est `null`.</span><span class="sxs-lookup"><span data-stu-id="f85c3-172">If multiple font colors are present within the range, `range.format.font.color` is `null`.</span></span>
 
### <a name="blank-input-for-a-property"></a><span data-ttu-id="f85c3-173">Entrée vide pour une propriété</span><span class="sxs-lookup"><span data-stu-id="f85c3-173">Blank input for a property</span></span>
 
<span data-ttu-id="f85c3-p120">Lorsque vous spécifiez une valeur vide pour une propriété (c’est-à-dire deux guillemets droits sans espace entre `''`), cela est interprété comme une instruction d’effacement ou de réinitialisation de la propriété. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="f85c3-p120">When you specify a blank value for a property (i.e., two quotation marks with no space in-between `''`), it will be interpreted as an instruction to clear or reset the property. For example:</span></span>
 
* <span data-ttu-id="f85c3-176">Si vous spécifiez une valeur vide pour la propriété `values` d’une plage, le contenu de la plage est effacé.</span><span class="sxs-lookup"><span data-stu-id="f85c3-176">If you specify a blank value for the `values` property of a range, the content of the range is cleared.</span></span>
 
* <span data-ttu-id="f85c3-177">Si vous spécifiez une valeur vide pour la propriété `numberFormat`, le format de nombre est réinitialisé sur `General`.</span><span class="sxs-lookup"><span data-stu-id="f85c3-177">If you specify a blank value for the `numberFormat` property, the number format is reset to `General`.</span></span>
 
* <span data-ttu-id="f85c3-178">Si vous spécifiez une valeur vide pour les propriétés `formula` et `formulaLocale`, les valeurs de la formule sont effacées.</span><span class="sxs-lookup"><span data-stu-id="f85c3-178">If you specify a blank value for the `formula` property and `formulaLocale` property, the formula values are cleared.</span></span>
 
### <a name="blank-property-values-in-the-response"></a><span data-ttu-id="f85c3-179">Valeurs de propriété vides dans la réponse</span><span class="sxs-lookup"><span data-stu-id="f85c3-179">Blank property values in the response</span></span>
 
<span data-ttu-id="f85c3-180">Pour les opérations de lecture, une valeur de propriété vide dans la réponse (c'est-à-dire, deux guillemets droits sans espace entre `''`) indique que la cellule ne contient pas de donnée ni de valeur.</span><span class="sxs-lookup"><span data-stu-id="f85c3-180">For read operations, a blank property value in the response (i.e., two quotation marks with no space in-between `''`) indicates that cell contains no data or value.</span></span> <span data-ttu-id="f85c3-181">Dans le premier exemple ci-dessous, la première et la dernière cellules de la plage ne contiennent pas de donnée.</span><span class="sxs-lookup"><span data-stu-id="f85c3-181">In the first example below, the first and last cell in the range contain no data.</span></span> <span data-ttu-id="f85c3-182">Dans le deuxième exemple, les deux premières cellules de la plage ne contiennent pas de formule.</span><span class="sxs-lookup"><span data-stu-id="f85c3-182">In the second example, the first two cells in the range do not contain a formula.</span></span>
 
```js
range.values = [['', 'some', 'data', 'in', 'other', 'cells', '']];
```
 
```js
range.formula = [['', '', '=Rand()']];
```
 
## <a name="read-or-write-to-an-unbounded-range"></a><span data-ttu-id="f85c3-183">Lire ou écrire dans une plage non liée</span><span class="sxs-lookup"><span data-stu-id="f85c3-183">Read or write to an unbounded range</span></span>
 
### <a name="read-an-unbounded-range"></a><span data-ttu-id="f85c3-184">Lire une plage non liée</span><span class="sxs-lookup"><span data-stu-id="f85c3-184">Read an unbounded range</span></span>
 
<span data-ttu-id="f85c3-p122">Une adresse de plage non liée est une adresse de plage qui spécifie des colonnes entières ou des lignes entières. Par exemple :</span><span class="sxs-lookup"><span data-stu-id="f85c3-p122">An unbounded range address is a range address that specifies either entire column(s) or entire row(s). For example:</span></span>
 
* <span data-ttu-id="f85c3-187">Adresses de plage composées de colonnes entières :</span><span class="sxs-lookup"><span data-stu-id="f85c3-187">Range addresses comprised of entire column(s):</span></span><ul><li>`C:C`</li><li>`A:F`</li></ul>
* <span data-ttu-id="f85c3-188">Adresses de plage composées de lignes entières :</span><span class="sxs-lookup"><span data-stu-id="f85c3-188">Range addresses comprised of entire row(s):</span></span><ul><li>`2:2`</li><li>`1:4`</li></ul>
 
<span data-ttu-id="f85c3-189">Lorsque l’API effectue une demande de récupération d’une plage non liée (par exemple, `getRange('C:C')`), la réponse contient des valeurs `null` pour les propriétés définies au niveau des cellules, telles que `values`, `text`, `numberFormat` et `formula`.</span><span class="sxs-lookup"><span data-stu-id="f85c3-189">When the API makes a request to retrieve an unbounded range (for example, `getRange('C:C')`), the response will contain `null` values for cell-level properties such as `values`, `text`, `numberFormat`, and `formula`.</span></span> <span data-ttu-id="f85c3-190">Les autres propriétés de la plage, telles que `address` et `cellCount`, contiennent des valeurs valides pour la plage non liée.</span><span class="sxs-lookup"><span data-stu-id="f85c3-190">Other properties of the range, such as `address` and `cellCount`, will contain valid values for the unbounded range.</span></span>
 
### <a name="write-to-an-unbounded-range"></a><span data-ttu-id="f85c3-191">Écrire dans une plage non liée</span><span class="sxs-lookup"><span data-stu-id="f85c3-191">Write to an unbounded range</span></span>
 
<span data-ttu-id="f85c3-192">Vous ne pouvez pas définir des propriétés au niveau de la cellule telles que `values`, `numberFormat`, et `formula` sur plage non liée, car la demande d’entrée  est trop volumineuse.</span><span class="sxs-lookup"><span data-stu-id="f85c3-192">You cannot set cell-level properties such as `values`, `numberFormat`, and `formula` on unbounded range because the input request is too large.</span></span> <span data-ttu-id="f85c3-193">Par exemple, l’extrait de code suivant n’est pas valide, car il tente de spécifier `values` pour une plage non liée.</span><span class="sxs-lookup"><span data-stu-id="f85c3-193">For example, the following code snippet is not valid because it attempts to specify `values` for an unbounded range.</span></span> <span data-ttu-id="f85c3-194">L’API renvoie une erreur si vous tentez de définir des propriétés au niveau de la cellule pour une plage non liée.</span><span class="sxs-lookup"><span data-stu-id="f85c3-194">The API will return an error if you attempt to set cell-level properties for an unbounded range.</span></span>
 
```js
const range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```
 
## <a name="read-or-write-to-a-large-range"></a><span data-ttu-id="f85c3-195">Lire ou écrire dans une grande plage</span><span class="sxs-lookup"><span data-stu-id="f85c3-195">Read or write to a large range</span></span>
 
<span data-ttu-id="f85c3-196">Si une plage contient un grand nombre de cellules, de valeurs, de formats de nombre et/ou de formules, il n’est peut-être pas possible d’exécuter des opérations d’API sur cette plage.</span><span class="sxs-lookup"><span data-stu-id="f85c3-196">If a range contains a large number of cells, values, number formats, and/or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="f85c3-197">L’API essaie toujours d’exécuter au mieux l’opération demandée sur une plage (par exemple, pour extraire ou écrire des données spécifiées), mais essayer d’effectuer des opérations de lecture ou d’écriture pour une grande plage peut provoquer une erreur d’API en raison de l’utilisation des ressources excessive.</span><span class="sxs-lookup"><span data-stu-id="f85c3-197">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="f85c3-198">Pour éviter ces erreurs, nous vous recommandons d’exécuter des opérations de lecture ou d’écriture distinctes pour des sous-ensembles plus petits d’une grande plage, au lieu d’essayer d’exécuter une seule opération de lecture ou d’écriture sur une grande plage.</span><span class="sxs-lookup"><span data-stu-id="f85c3-198">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>
 
## <a name="update-all-cells-in-a-range"></a><span data-ttu-id="f85c3-199">Mettre à jour toutes les cellules d’une plage</span><span class="sxs-lookup"><span data-stu-id="f85c3-199">Update all cells in a range</span></span>
 
<span data-ttu-id="f85c3-200">Pour appliquer la même mise à jour à toutes les cellules d’une plage, (par exemple, pour remplir toutes les cellules avec la même valeur, définir le même format de nombre ou renseigner toutes les cellules avec la même formule), définissez la propriété correspondante dans l’objet **range** sur la valeur (unique) de votre choix.</span><span class="sxs-lookup"><span data-stu-id="f85c3-200">To apply the same update to all cells in a range, (for example, to populate all cells with the same value, set the same number format, or populate all cells with the same formula), set the corresponding property on the **range** object to the desired (single) value.</span></span>
 
<span data-ttu-id="f85c3-201">L’exemple suivant obtient une plage qui contient 20 cellules, puis définit le format de nombre et remplit toutes les cellules de la plage avec la valeur **3/11/2015**.</span><span class="sxs-lookup"><span data-stu-id="f85c3-201">The following example gets a range that contains 20 cells, and then sets the number format and populates all cells in the range with the value **3/11/2015**.</span></span>
 
```js
Excel.run(function (context) {
  const sheetName = 'Sheet1';
  const rangeAddress = 'A1:A20';
  const worksheet = context.workbook.worksheets.getItem(sheetName);
 
  const range = worksheet.getRange(rangeAddress);
  range.numberFormat = 'm/d/yyyy';
  range.values = '3/11/2015';
  range.load('text');
 
  return context.sync()
    .then(function () {
      console.log(range.text);
  });
}).catch(function (error) {
  console.log('Error: ' + error);
  if (error instanceof OfficeExtension.Error) {
    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
  }
});
```
 
##<a name="handle-errors"></a><span data-ttu-id="f85c3-202">Gestion des erreurs</span><span class="sxs-lookup"><span data-stu-id="f85c3-202">Handle errors</span></span>

<span data-ttu-id="f85c3-203">Lorsqu’une erreur d’API se produit, l’API renvoie un objet **error** qui contient un code et un message.</span><span class="sxs-lookup"><span data-stu-id="f85c3-203">When an API error occurs, the API returns an **error** object that contains a code and a message.</span></span> <span data-ttu-id="f85c3-204">Pour plus d’informations sur la gestion des erreurs, notamment la liste des erreurs d’API, consultez la rubrique [Gestion des erreurs](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="f85c3-204">For detailed information about error handling, including a list of API errors, see [Error handling](excel-add-ins-error-handling.md).</span></span>
 
## <a name="see-also"></a><span data-ttu-id="f85c3-205">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="f85c3-205">See also</span></span>
 
* [<span data-ttu-id="f85c3-206">Prise en main des compléments Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-206">Get started with Excel add-ins</span></span>](excel-add-ins-get-started-overview.md)
* [<span data-ttu-id="f85c3-207">Exemples de code pour les compléments Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-207">Excel add-ins code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples)
* [<span data-ttu-id="f85c3-208">Concepts avancés de programmation avec l’API JavaScript Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-208">Advanced programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-advanced-concepts.md)
* [<span data-ttu-id="f85c3-209">Optimisation des performances à l’aide de l’API JavaScript d’Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-209">Excel JavaScript API performance optimization</span></span>](https://docs.microsoft.com/office/dev/add-ins/excel/performance)
* [<span data-ttu-id="f85c3-210">Référence de l’API JavaScript pour Excel</span><span class="sxs-lookup"><span data-stu-id="f85c3-210">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
