---
title: Concepts fondamentaux de programmation avec l’API JavaScript pour Word
description: L’API JavaScript pour Word permet de créer des compléments pour Word.
ms.date: 07/05/2019
localization_priority: Priority
ms.openlocfilehash: 7849780c1aed48152355c3fdbf350d798b2de1f2
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42325016"
---
# <a name="fundamental-programming-concepts-with-the-word-javascript-api"></a><span data-ttu-id="dc9c4-103">Concepts fondamentaux de programmation avec l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-103">Fundamental programming concepts with the Word JavaScript API</span></span>

<span data-ttu-id="dc9c4-104">Cet article décrit les concepts de base de l’utilisation de [l’API JavaScript pour Word](../reference/overview/word-add-ins-reference-overview.md) pour créer des compléments pour Word 2016 ou version ultérieure.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-104">This article describes concepts that are fundamental to using the [Word JavaScript API](../reference/overview/word-add-ins-reference-overview.md) to build add-ins for Word 2016 or later.</span></span>

## <a name="referencing-officejs"></a><span data-ttu-id="dc9c4-105">Referencing Office.js</span><span class="sxs-lookup"><span data-stu-id="dc9c4-105">Referencing Office.js</span></span>

<span data-ttu-id="dc9c4-106">Vous pouvez référencer Office.js à partir des emplacements suivants :</span><span class="sxs-lookup"><span data-stu-id="dc9c4-106">You can reference Office.js from the following locations:</span></span>

- <span data-ttu-id="dc9c4-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` : utilisez cette ressource pour les compléments de production.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-107">`https://appsforoffice.microsoft.com/lib/1/hosted/office.js` - use this resource for production add-ins.</span></span>

- <span data-ttu-id="dc9c4-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` : utilisez cette ressource pour essayer les fonctionnalités en préversion.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-108">`https://appsforoffice.microsoft.com/lib/beta/hosted/office.js` - use this resource to try out preview features.</span></span>

## <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="dc9c4-109">Ensembles de conditions requises de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-109">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="dc9c4-110">Les ensembles de conditions requises sont des groupes nommés de membres d’API.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-110">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="dc9c4-111">Les compléments Office utilisent les ensembles de conditions requises spécifiés dans le manifeste ou utilisent une vérification de l’exécution pour déterminer si un hôte Office prend en charge les API requises par le complément.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-111">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs.</span></span> <span data-ttu-id="dc9c4-112">Pour en savoir plus sur les ensembles de conditions requises de l’API JavaScript pour Word, voir [Ensembles de conditions requises de l’API JavaScript pour Word](../reference/requirement-sets/word-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="dc9c4-112">For detailed information about Word JavaScript API requirement sets, see [Word JavaScript API requirement sets](../reference/requirement-sets/word-api-requirement-sets.md).</span></span>

## <a name="running-word-add-ins"></a><span data-ttu-id="dc9c4-113">Exécution de compléments Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-113">Running Word add-ins</span></span>

<span data-ttu-id="dc9c4-114">Pour exécuter votre complément, utilisez un gestionnaire d’événements `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-114">To run your add-in, use an `Office.initialize` event handler.</span></span> <span data-ttu-id="dc9c4-115">Pour plus d’informations sur l’initialisation du complément, voir [Présentation de l’API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span><span class="sxs-lookup"><span data-stu-id="dc9c4-115">For more information about add-in initialization, see [Understanding the API](/office/dev/add-ins/develop/understanding-the-javascript-api-for-office).</span></span>

<span data-ttu-id="dc9c4-116">Les compléments qui ciblent Word 2016 ou version ultérieure s’exécutent en transmettant une fonction dans la méthode `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-116">Add-ins that target Word 2016 or later run by passing a function into the `Word.run()` method.</span></span> <span data-ttu-id="dc9c4-117">La fonction transmise dans la méthode `run` doit contenir un argument de contexte.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-117">The function passed into the `run` method must have a context argument.</span></span> <span data-ttu-id="dc9c4-118">Cet [objet de contexte](/javascript/api/word/word.requestcontext) est différent de celui que vous obtenez de l’objet Office, même s’il sert également à interagir avec l’environnement d’exécution de Word.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-118">This [context object](/javascript/api/word/word.requestcontext) is different than the context object you get from the Office object, but it is also used to interact with the Word runtime environment.</span></span> <span data-ttu-id="dc9c4-119">L’objet de contexte permet d’accéder au modèle objet de l’API JavaScript pour Word.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-119">The context object provides access to the Word JavaScript API object model.</span></span> <span data-ttu-id="dc9c4-120">L’exemple suivant montre comment initialiser et exécuter un complément Word à l’aide de la méthode `Word.run()`.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-120">The following example shows how to initialize and run a Word add-in by using the `Word.run()` method.</span></span>

```js
(function () {
    "use strict";

    // The initialize event handler must be run on each page to initialize Office JS.
    // You can add optional custom initialization code that will run after OfficeJS
    // has initialized.
    Office.initialize = function (reason) {
        // The reason object tells how the add-in was initialized. The values can be:
        // inserted - the add-in was inserted to an open document.
        // documentOpened - the add-in was already inserted in to the document and the document was opened.

        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // Set your optional initialization code.
            // You can also load saved settings from the Office object.
        });
    };

    // Run a batch operation against the Word JavaScript API object model.
    // Use the context argument to get access to the Word document.
    Word.run(function (context) {

        // Create a proxy object for the document.
        var thisDocument = context.document;
        // ...
    })
})();
```

### <a name="asynchronous-nature-of-word-apis"></a><span data-ttu-id="dc9c4-121">Nature asynchrone des API pour Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-121">Asynchronous nature of Word APIs</span></span>

<span data-ttu-id="dc9c4-122">L’API JavaScript pour Word est chargée par Office.js.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-122">The Word JavaScript API is loaded by Office.js.</span></span> <span data-ttu-id="dc9c4-123">L’API JavaScript pour Word change la façon d’interagir avec des objets tels que des documents et des paragraphes.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-123">The Word JavaScript API changes the way that you can interact with objects like documents and paragraphs.</span></span> <span data-ttu-id="dc9c4-124">Ainsi, vous ne devez plus utiliser d’API asynchrones individuelles pour extraire et mettre à jour chacun de ces objets. L’API JavaScript pour Word fournit des objets JavaScript « proxy » qui correspondent aux objets en direct s’exécutant dans Word.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-124">Rather than providing individual asynchronous APIs for retrieving and updating each of these objects, the Word JavaScript API provides "proxy" JavaScript objects that correspond to the live objects running in Word.</span></span> <span data-ttu-id="dc9c4-125">Vous pouvez interagir avec ces objets proxy en lisant et écrivant leurs propriétés de façon synchronisée, et en appelant des méthodes synchrones pour effectuer des opérations sur ces objets.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-125">You can interact with these proxy objects by synchronously reading and writing their properties and calling synchronous methods to perform operations on them.</span></span> <span data-ttu-id="dc9c4-126">Ces interactions avec des objets proxy n’ont pas lieu immédiatement dans le script en cours d’exécution.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-126">These interactions with proxy objects aren't immediately realized in the running script.</span></span> <span data-ttu-id="dc9c4-127">La méthode `context.sync` synchronise l’état de vos objets JavaScript en cours d’exécution et celui des objets réels en exécutant des instructions en file d’attente et en récupérant des propriétés d’objets Word chargés pour les utiliser dans votre script.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-127">The `context.sync` method synchronizes the state between your running JavaScript and the real objects in Office by executing queued instructions and retrieving properties of loaded Word objects for use in your script.</span></span>

### <a name="synchronizing-word-documents-with-word-javascript-api-proxy-objects"></a><span data-ttu-id="dc9c4-128">Synchronisation de documents Word avec des objets de proxy de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-128">Synchronizing Word documents with Word JavaScript API proxy objects</span></span>

<span data-ttu-id="dc9c4-p105">Le modèle objet de l’API JavaScript pour Word est associé de façon relativement libre aux objets dans Word. Les objets de l’API JavaScript pour Word sont des proxys pour des objets dans un document Word. Les actions effectuées sur les objets de proxy ne sont pas réalisées dans Word tant que l’état du document n’a pas été synchronisé. Inversement, l’état du document Word n’est pas répercuté sur les objets de proxy tant que l’état du document n’a pas été synchronisé. Pour synchroniser l’état du document, vous exécutez la méthode `context.sync()`. L’exemple suivant présente la création d’un objet Body de proxy et une file de commandes permettant de charger la propriété de texte sur l’objet Body de proxy, puis la synchronisation du corps dans le document Word avec l’objet de proxy correspondant à l’aide de la méthode `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-p105">The Word JavaScript API object model is loosely coupled with the objects in Word. Word JavaScript API objects are proxies for objects in a Word document. Actions taken on proxy objects are not realized in Word until the document state has been synchronized. Conversely, the state of the Word document is not realized in the proxy objects until the document state has been synchronized. To synchronize the document state, you run the `context.sync()` method. The following example creates a proxy body object and a queued command to load the text property on the proxy body object, and uses the `context.sync()` method to synchronize the body of the Word document with the body proxy object.</span></span>

```js
// Run a batch operation against the Word object model.
Word.run(function (context) {

    // Create a proxy object for the document body.
    // The body object hasn't been set with any property values.
    var body = context.document.body;

    // Queue a command to load the text property for the proxy document body object.
    body.load("text");

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

### <a name="executing-a-batch-of-commands"></a><span data-ttu-id="dc9c4-135">Exécution d’un lot de commandes</span><span class="sxs-lookup"><span data-stu-id="dc9c4-135">Executing a batch of commands</span></span>

<span data-ttu-id="dc9c4-136">Les objets de proxy Word utilisent des méthodes pour accéder au modèle objet et le mettre à jour.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-136">The Word proxy objects have methods for accessing and updating the object model.</span></span> <span data-ttu-id="dc9c4-137">Ces méthodes sont exécutées séquentiellement, dans l’ordre de leur mise en file d’attente dans le lot.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-137">These methods are run sequentially in the order in which they were queued in the batch.</span></span> <span data-ttu-id="dc9c4-138">Toutes les commandes en file d’attente dans le lot sont exécutées lors de l’appel de la méthode `context.sync()`.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-138">All of the commands that are queued in the batch are run when `context.sync()` is called.</span></span>

<span data-ttu-id="dc9c4-139">L’exemple suivant montre comment fonctionne la file d’attente de commandes.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-139">The following example shows how the command queue works.</span></span> <span data-ttu-id="dc9c4-140">Lors de l’appel de la méthode `context.sync()`, la commande visant à charger le corps du texte est exécutée dans Word.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-140">When `context.sync()` is called, the command to load the body text is run in Word.</span></span> <span data-ttu-id="dc9c4-141">C’est ensuite la commande visant à insérer du texte dans le corps de Word qui est appliquée.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-141">Then, the command to insert text into the body in Word occurs.</span></span> <span data-ttu-id="dc9c4-142">Les résultats sont alors renvoyés vers l’objet Body de proxy.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-142">The results are then returned to the body proxy object.</span></span> <span data-ttu-id="dc9c4-143">La valeur de la propriété `body.text` dans le code JavaScript Word est la même que celle du corps du document de Word <u>avant</u> l’insertion du texte dans le document Word.</span><span class="sxs-lookup"><span data-stu-id="dc9c4-143">The value of the `body.text` property in the Word JavaScript API is the value of the Word document body <u>before</u> the text was inserted into Word document.</span></span>

```js
// Run a batch operation against the Word JavaScript API.
Word.run(function (context) {

    // Create a proxy object for the document body.
    var body = context.document.body;

    // Queue a command to load the text property of the proxy body object.
    body.load("text");

    // Queue a command to insert text into the end of the Word document body.
    body.insertText('This is text inserted after loading the body.text property',
                    Word.InsertLocation.end);

    // Synchronize the document state by executing the queued commands,
    // and return a promise to indicate task completion.
    return context.sync().then(function () {
        console.log("Body contents: " + body.text);
    });
})
```

## <a name="see-also"></a><span data-ttu-id="dc9c4-144">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="dc9c4-144">See also</span></span>

- [<span data-ttu-id="dc9c4-145">Présentation de l’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-145">Word JavaScript API overview</span></span>](../reference/overview/word-add-ins-reference-overview.md)
- [<span data-ttu-id="dc9c4-146">Créer votre premier complément Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-146">Build your first Word add-in</span></span>](../quickstarts/word-quickstart.md)
- [<span data-ttu-id="dc9c4-147">Didacticiel sur les compléments Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-147">Word add-in tutorial</span></span>](../tutorials/word-tutorial.md)
- [<span data-ttu-id="dc9c4-148">Référence d’API JavaScript pour Word</span><span class="sxs-lookup"><span data-stu-id="dc9c4-148">Word JavaScript API reference</span></span>](/javascript/api/word)