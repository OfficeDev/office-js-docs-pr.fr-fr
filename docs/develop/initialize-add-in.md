---
title: Initialiser votre complément Office
description: Découvrez comment initialiser votre complément Office.
ms.date: 02/27/2020
localization_priority: Normal
ms.openlocfilehash: 5dc9d0143ac9eaab18625e280891bd601fa9f899
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293323"
---
# <a name="initialize-your-office-add-in"></a><span data-ttu-id="198f3-103">Initialiser votre complément Office</span><span class="sxs-lookup"><span data-stu-id="198f3-103">Initialize your Office Add-in</span></span>

<span data-ttu-id="198f3-104">Les compléments Office ont souvent une logique de démarrage pour effectuer des actions telles que :</span><span class="sxs-lookup"><span data-stu-id="198f3-104">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="198f3-105">Vérifiez que la version de l’utilisateur d’Office prend en charge toutes les API Office que votre code appelle.</span><span class="sxs-lookup"><span data-stu-id="198f3-105">Check that the user's version of Office supports all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="198f3-106">Vérifier l’existence de certains artefacts, tels qu’une feuille de calcul avec un nom spécifique.</span><span class="sxs-lookup"><span data-stu-id="198f3-106">Ensure the existence of certain artifacts, such as a worksheet with a specific name.</span></span>

- <span data-ttu-id="198f3-107">Inviter l’utilisateur à sélectionner certaines cellules dans Excel, puis insérer un graphique initialisé avec ces valeurs sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="198f3-107">Prompt the user to select some cells in Excel, and then insert a chart initialized with those selected values.</span></span>

- <span data-ttu-id="198f3-108">Établir des liaisons.</span><span class="sxs-lookup"><span data-stu-id="198f3-108">Establish bindings.</span></span>

- <span data-ttu-id="198f3-109">Utiliser l’API de boîte de dialogue Office pour inviter l’utilisateur à entrer les valeurs des paramètres de complément par défaut.</span><span class="sxs-lookup"><span data-stu-id="198f3-109">Use the Office Dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="198f3-110">Toutefois, un complément Office ne peut pas appeler d’API JavaScript Office tant que la bibliothèque n’a pas été chargée.</span><span class="sxs-lookup"><span data-stu-id="198f3-110">However, an Office Add-in cannot successfully call any Office JavaScript APIs until the library has been loaded.</span></span> <span data-ttu-id="198f3-111">Cet article décrit les deux façons dont votre code peut s’assurer que la bibliothèque a été chargée :</span><span class="sxs-lookup"><span data-stu-id="198f3-111">This article describes the two ways your code can ensure that the library has been loaded:</span></span>

- <span data-ttu-id="198f3-112">Initialiser avec `Office.onReady()` .</span><span class="sxs-lookup"><span data-stu-id="198f3-112">Initialize with `Office.onReady()`.</span></span>
- <span data-ttu-id="198f3-113">Initialiser avec `Office.initialize` .</span><span class="sxs-lookup"><span data-stu-id="198f3-113">Initialize with `Office.initialize`.</span></span>

> [!TIP]
> <span data-ttu-id="198f3-114">Au lieu de `Office.initialize`, nous vous recommandons d’utiliser `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="198f3-114">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="198f3-115">Bien que `Office.initialize` est toujours pris en charge, `Office.onReady()` offre davantage de flexibilité.</span><span class="sxs-lookup"><span data-stu-id="198f3-115">Although `Office.initialize` is still supported, `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="198f3-116">Vous ne pouvez attribuer qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure Office.</span><span class="sxs-lookup"><span data-stu-id="198f3-116">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure.</span></span> <span data-ttu-id="198f3-117">Vous pouvez appeler `Office.onReady()` à différents endroits de votre code et utiliser des rappels différents.</span><span class="sxs-lookup"><span data-stu-id="198f3-117">You can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="198f3-118">Pour plus d’informations sur les différences entre ces techniques, reportez-vous à la rubrique [Différences majeures entre Office.initialize et Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="198f3-118">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="198f3-119">Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et environnement d’exécution](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="198f3-119">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

## <a name="initialize-with-officeonready"></a><span data-ttu-id="198f3-120">Initialiser avec Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="198f3-120">Initialize with Office.onReady()</span></span>

<span data-ttu-id="198f3-121">`Office.onReady()` est une méthode asynchrone qui renvoie un objet [promesse](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) pendant qu’il vérifie si la bibliothèque Office.js est chargée.</span><span class="sxs-lookup"><span data-stu-id="198f3-121">`Office.onReady()` is an asynchronous method that returns a [Promise](https://developer.mozilla.org/docs/Web/JavaScript/Reference/Global_Objects/Promise) object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="198f3-122">Lorsque la bibliothèque est chargée, elle résout la promesse comme un objet qui spécifie l’application cliente Office avec une `Office.HostType` valeur enum ( `Excel` , `Word` , etc.) et la plateforme avec une `Office.PlatformType` valeur enum (,, `PC` `Mac` `OfficeOnline` , etc.).</span><span class="sxs-lookup"><span data-stu-id="198f3-122">When the library is loaded, it resolves the Promise as an object that specifies the Office client application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="198f3-123">L’objet Promise se résout immédiatement si la bibliothèque est déjà chargée quand `Office.onReady()` est appelée.</span><span class="sxs-lookup"><span data-stu-id="198f3-123">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="198f3-124">Une méthode pour appeler `Office.onReady()` consiste à transmettre une méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="198f3-124">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="198f3-125">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="198f3-125">Here's an example:</span></span>

```js
Office.onReady(function(info) {
    if (info.host === Office.HostType.Excel) {
        // Do Excel-specific initialization (for example, make add-in task pane's
        // appearance compatible with Excel "green").
    }
    if (info.platform === Office.PlatformType.PC) {
        // Make minor layout changes in the task pane.
    }
    console.log(`Office.js is now ready in ${info.host} on ${info.platform}`);
});
```

<span data-ttu-id="198f3-126">Par ailleurs, vous pouvez mettre en chaîne une`then()` méthode permettant d’appeler `Office.onReady()`, au lieu de spécifier un rappel.</span><span class="sxs-lookup"><span data-stu-id="198f3-126">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="198f3-127">Par exemple, le code suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.</span><span class="sxs-lookup"><span data-stu-id="198f3-127">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="198f3-128">Voici le même exemple utilisant les mots clés `async` et `await` dans TypeScript :</span><span class="sxs-lookup"><span data-stu-id="198f3-128">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="198f3-129">Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être*habituellement* placées dans la réponse à`Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="198f3-129">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="198f3-130">Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :</span><span class="sxs-lookup"><span data-stu-id="198f3-130">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="198f3-131">Toutefois, il existe des exceptions à cette pratique.</span><span class="sxs-lookup"><span data-stu-id="198f3-131">However, there are exceptions to this practice.</span></span> <span data-ttu-id="198f3-132">Par exemple, supposons que vous souhaitez ouvrir votre complément dans un navigateur (au lieu de le chargement dans une application Office) afin de déboguer votre interface utilisateur à l’aide des outils de navigation.</span><span class="sxs-lookup"><span data-stu-id="198f3-132">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office application) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="198f3-133">Étant donné que Office.js ne sera pas chargé dans le navigateur, `onReady` ne s’exécutera pas et le `$(document).ready` ne s’exécutera pas si cette opération est appelée à l’intérieur d’Office `onReady`.</span><span class="sxs-lookup"><span data-stu-id="198f3-133">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> 

<span data-ttu-id="198f3-134">Il est également possible d’afficher un indicateur de progression dans le volet Office pendant le chargement du complément.</span><span class="sxs-lookup"><span data-stu-id="198f3-134">Another exception would be if you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="198f3-135">Dans ce scénario, votre code doit appeler jQuery `ready` et utiliser son rappel pour afficher l’indicateur de progression.</span><span class="sxs-lookup"><span data-stu-id="198f3-135">In this scenario, your code should call the jQuery `ready` and use its callback to render the progress indicator.</span></span> <span data-ttu-id="198f3-136">Puis le rappel `onReady` Office peut remplacer l’indicateur de progression par l’interface utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="198f3-136">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

## <a name="initialize-with-officeinitialize"></a><span data-ttu-id="198f3-137">Initialiser avec Office.initialize</span><span class="sxs-lookup"><span data-stu-id="198f3-137">Initialize with Office.initialize</span></span>

<span data-ttu-id="198f3-138">Un événement initialisé se déclenche lorsque la bibliothèque Office.js est chargée et prête pour une interaction avec l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="198f3-138">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="198f3-139">Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente votre logique d’initialisation.</span><span class="sxs-lookup"><span data-stu-id="198f3-139">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="198f3-140">L’exemple suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.</span><span class="sxs-lookup"><span data-stu-id="198f3-140">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="198f3-141">Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, ceux-ci doivent *généralement* être placés au sein de l' `Office.initialize` événement (les exceptions décrites dans la section **Initialize with Office. onReady ()** ci-dessus s’appliquent également dans ce cas).</span><span class="sxs-lookup"><span data-stu-id="198f3-141">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event (the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also).</span></span> <span data-ttu-id="198f3-142">Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :</span><span class="sxs-lookup"><span data-stu-id="198f3-142">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="198f3-143">Pour les compléments de tâches et de contenu, `Office.initialize` fournit un paramètre_raison_ supplémentaire.</span><span class="sxs-lookup"><span data-stu-id="198f3-143">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="198f3-144">Ce paramètre peut être utilisé pour savoir comment un complément a été ajouté au document actif.</span><span class="sxs-lookup"><span data-stu-id="198f3-144">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="198f3-145">Vous pouvez l’utiliser pour fournir une logique différente quand un complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document.</span><span class="sxs-lookup"><span data-stu-id="198f3-145">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

```js
Office.initialize = function (reason) {
    $(document).ready(function () {
        switch (reason) {
            case 'inserted': console.log('The add-in was just inserted.');
            case 'documentOpened': console.log('The add-in is already part of the document.');
        }
    });
 };
```

<span data-ttu-id="198f3-146">Pour plus d’informations, consultez les pages relatives à l’[événement Office.initialize](/javascript/api/office) et à l’[énumération InitializationReason](/javascript/api/office/office.initializationreason).</span><span class="sxs-lookup"><span data-stu-id="198f3-146">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

## <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="198f3-147">Principales différences entre Office.initialize et Office.onReady</span><span class="sxs-lookup"><span data-stu-id="198f3-147">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="198f3-148">Vous ne pouvez assigner qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure d’Office, mais vous pouvez appeler `Office.onReady()`à plusieurs endroits dans votre code et utiliser des rappels différents.</span><span class="sxs-lookup"><span data-stu-id="198f3-148">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="198f3-149">Par exemple, votre code pourrait appeler `Office.onReady()` dès que votre script personnalisé charge avec un rappel qui exécute la logique d’initialisation ; et votre code peut également comporter un bouton dans le volet Office dont le script appelle `Office.onReady()` avec un rappel différent.</span><span class="sxs-lookup"><span data-stu-id="198f3-149">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="198f3-150">Si c’est le cas, le deuxième rappel s’exécute quand l’utilisateur clique sur le bouton.</span><span class="sxs-lookup"><span data-stu-id="198f3-150">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="198f3-151">L’événement`Office.initialize` se déclenche à la fin du processus interne dans lequel Office.js s’initialise lui-même.</span><span class="sxs-lookup"><span data-stu-id="198f3-151">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="198f3-152">Et il se déclenche *immédiatement* après la fin du processus interne.</span><span class="sxs-lookup"><span data-stu-id="198f3-152">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="198f3-153">Si le code dans lequel vous attribuez un gestionnaire à l’événement s’exécute trop longtemps après le déclenchement de l’événement, votre gestionnaire ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="198f3-153">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="198f3-154">Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du complément pour charger les fichiers polyfill une fois que le serveur charge Office.js mais avant que le serveur ne charge votre code JavaScript personnalisé.</span><span class="sxs-lookup"><span data-stu-id="198f3-154">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="198f3-155">Le temps que votre script se charge et affecte le Gestionnaire, l’événement initialisé s’est déjà produit.</span><span class="sxs-lookup"><span data-stu-id="198f3-155">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="198f3-156">Mais il n’est jamais « trop tard » pour appeler `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="198f3-156">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="198f3-157">Si l’événement initialisé s’est déjà produit, le rappel s’exécute immédiatement.</span><span class="sxs-lookup"><span data-stu-id="198f3-157">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="198f3-158">Même si vous n’avez aucune logique de démarrage, appelez `Office.onReady()` ou attribuez une fonction vide à `Office.initialize` lorsque votre complément JavaScript se charge.</span><span class="sxs-lookup"><span data-stu-id="198f3-158">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="198f3-159">Certaines combinaisons d’applications Office et de plateforme ne chargent pas le volet des tâches tant que l’un des événements suivants se produit.</span><span class="sxs-lookup"><span data-stu-id="198f3-159">Some Office application and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="198f3-160">Les exemples suivants présentent ces deux approches.</span><span class="sxs-lookup"><span data-stu-id="198f3-160">The following examples show these two approaches.</span></span>
>
>```js    
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="see-also"></a><span data-ttu-id="198f3-161">Voir aussi</span><span class="sxs-lookup"><span data-stu-id="198f3-161">See also</span></span>

- [<span data-ttu-id="198f3-162">Compréhension de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="198f3-162">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="198f3-163">Chargement du DOM et de l’environnement d’exécution</span><span class="sxs-lookup"><span data-stu-id="198f3-163">Loading the DOM and runtime environment</span></span>](loading-the-dom-and-runtime-environment.md)