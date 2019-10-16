---
title: Présentation de l’API JavaScript pour Office
description: ''
ms.date: 06/21/2019
localization_priority: Priority
ms.openlocfilehash: 1954457b477472b8940841bb1ffe5954e49e01ec
ms.sourcegitcommit: b3996b1444e520b44cf752e76eef50908386ca26
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 10/16/2019
ms.locfileid: "37524233"
---
# <a name="understanding-the-javascript-api-for-office"></a><span data-ttu-id="9e7f9-102">Présentation de l’API JavaScript pour Office</span><span class="sxs-lookup"><span data-stu-id="9e7f9-102">Understanding the JavaScript API for Office</span></span>

<span data-ttu-id="9e7f9-p101">Cet article fournit des informations sur l’API JavaScript pour Office et son utilisation. Pour obtenir des informations de référence, voir [API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office). Pour plus d’informations sur la mise à jour des fichiers de projet Visual Studio vers la version la plus récente de l’API JavaScript pour Office, voir [Mettre à jour la version de votre API JavaScript pour Office et les fichiers de schéma manifeste](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-p101">This article provides information about the JavaScript API for Office and how to use it. For reference information, see [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office). For information about updating Visual Studio project files to the most current version of the JavaScript API for Office, see [Update the version of your JavaScript API for Office and manifest schema files](update-your-javascript-api-for-office-and-manifest-schema-version.md).</span></span>

> [!NOTE]
> <span data-ttu-id="9e7f9-p102">Si vous prévoyez de [publier](../publish/publish.md) votre complément sur AppSource et de le rendre disponible dans l’expérience Office, assurez-vous que vous respectez les [stratégies de validation AppSource](/office/dev/store/validation-policies). Par exemple, pour réussir la validation, votre complément doit fonctionner sur toutes les plateformes prenant en charge les méthodes définies (pour en savoir plus, consultez la [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) et la [page relative à la disponibilité des compléments Office sur les plateformes et les hôtes](../overview/office-add-in-availability.md)).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-p102">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

## <a name="referencing-the-javascript-api-for-office-library-in-your-add-in"></a><span data-ttu-id="9e7f9-108">Référencer la bibliothèque de l’interface API JavaScript pour Office dans votre complément</span><span class="sxs-lookup"><span data-stu-id="9e7f9-108">Referencing the JavaScript API for Office library in your add-in</span></span>

<span data-ttu-id="9e7f9-p103">La bibliothèque de l’[interface API JavaScript pour Office](/office/dev/add-ins/reference/javascript-api-for-office) comprend le fichier Office.js et des fichiers .js propres aux applications hôtes associées, comme Excel-15.js et Outlook15.js. La méthode la plus simple pour référencer l’interface API est d’utiliser notre CDN en ajoutant le `<script>` suivant à la balise `<head>` de votre page :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-p103">The [JavaScript API for Office](/office/dev/add-ins/reference/javascript-api-for-office) library consists of the Office.js file and associated host application-specific .js files, such as Excel-15.js and Outlook-15.js. The simplest method of referencing the API is using our CDN by adding the following `<script>` to your page's `<head>` tag:</span></span>  

```html
<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
```

<span data-ttu-id="9e7f9-111">Cette opération permet de télécharger et de mettre en cache les fichiers de l’interface API JavaScript pour Office lors du premier chargement de votre complément pour garantir qu’elle utilise l’implémentation d’Office.js la plus récente et les fichiers .js qui lui sont associés pour la version indiquée.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-111">This will download and cache the JavaScript API for Office files the first time your add-in loads to make sure that it is using the most up-to-date implementation of Office.js and its associated files for the specified version.</span></span>

<span data-ttu-id="9e7f9-112">Pour obtenir plus d’informations sur le CDN Office.js et la gestion du contrôle de version et de la rétrocompatibilité, consultez la page relative au [référencement de la bibliothèque de l’interface API JavaScript pour Office à partir de son réseau de distribution de contenu (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-112">For more details around the Office.js CDN, including how versioning and backward compatibility is handled, see [Referencing the JavaScript API for Office library from its content delivery network (CDN)](referencing-the-javascript-api-for-office-library-from-its-cdn.md).</span></span>

## <a name="initializing-your-add-in"></a><span data-ttu-id="9e7f9-113">Initialisation de votre complément</span><span class="sxs-lookup"><span data-stu-id="9e7f9-113">Initializing your add-in</span></span>

<span data-ttu-id="9e7f9-114">**S’applique à :** tous les types de complément</span><span class="sxs-lookup"><span data-stu-id="9e7f9-114">**Applies to:** All add-in types</span></span>

<span data-ttu-id="9e7f9-115">Les compléments Office ont souvent une logique de démarrage pour effectuer des actions telles que :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-115">Office Add-ins often have start-up logic to do things such as:</span></span>

- <span data-ttu-id="9e7f9-116">Vérifiez que version de l’utilisateur d’Office prendra en charge tous les API Office que votre code appelle.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-116">Check that the user's version of Office will support all the Office APIs that your code calls.</span></span>

- <span data-ttu-id="9e7f9-117">Vérifiez l’existence de certains artefacts tels que des feuille de calcul avec un nom spécifique.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-117">Ensure the existence of certain artifacts, such as worksheet with a specific name.</span></span>

- <span data-ttu-id="9e7f9-118">Avertir l’utilisateur pour sélectionner certaines cellules dans Excel, puis insérer un graphique initialisé avec ces valeurs sélectionnées.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-118">Prompting the user to select some cells in Excel, and then inserting a chart initialized with those selected values.</span></span>

- <span data-ttu-id="9e7f9-119">Établir des liaisons.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-119">Establish bindings.</span></span>

- <span data-ttu-id="9e7f9-120">Utilisez la boîte de dialogue Office API pour inviter l’utilisateur pour les valeurs de paramètres des compléments par défaut.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-120">Use the Office dialog API to prompt the user for default add-in settings values.</span></span>

<span data-ttu-id="9e7f9-121">Mais votre code de démarrage ne doit pas appeler une API Office.js tant que la bibliothèque n’est pas chargée.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-121">But your start-up code must not call any Office.js APIs until the library is loaded.</span></span> <span data-ttu-id="9e7f9-122">Il existe deux manières pour votre code de s’assurer que la bibliothèque est chargée.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-122">There are two ways that your code can ensure that the library is loaded.</span></span> <span data-ttu-id="9e7f9-123">Ceci est décrit en détail dans les sections ci-après :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-123">They are described in the following sections:</span></span> 

- [<span data-ttu-id="9e7f9-124">Initialiser avec Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="9e7f9-124">Initialize with Office.onReady()</span></span>](#initialize-with-officeonready)
- [<span data-ttu-id="9e7f9-125">Initialiser avec Office.initialize</span><span class="sxs-lookup"><span data-stu-id="9e7f9-125">Initialize with Office.initialize</span></span>](#initialize-with-officeinitialize)

> [!TIP]
> <span data-ttu-id="9e7f9-126">Au lieu de `Office.initialize`, nous vous recommandons d’utiliser `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-126">We recommend that you use `Office.onReady()` instead of `Office.initialize`.</span></span> <span data-ttu-id="9e7f9-127">`Office.initialize` est toujours pris en charge, mais `Office.onReady()` offre plus de flexibilité.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-127">Although `Office.initialize` is still supported, using `Office.onReady()` provides more flexibility.</span></span> <span data-ttu-id="9e7f9-128">Vous ne pouvez assigner qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure d’Office, mais vous pouvez appeler `Office.onReady()`à plusieurs endroits dans votre code et utiliser des rappels différents.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-128">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure, but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span>
> 
> <span data-ttu-id="9e7f9-129">Pour plus d’informations sur les différences entre ces techniques, reportez-vous à la rubrique [Différences majeures entre Office.initialize et Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-129">For information about the differences in these techniques, see [Major differences between Office.initialize and Office.onReady()](#major-differences-between-officeinitialize-and-officeonready).</span></span>

<span data-ttu-id="9e7f9-130">Pour plus de détails sur la séquence d’événements lors de l’initialisation d’un complément, reportez-vous à la rubrique [Chargement du DOM et environnement d’exécution](loading-the-dom-and-runtime-environment.md).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-130">For more details about the sequence of events when an add-in is initialized, see [Loading the DOM and runtime environment](loading-the-dom-and-runtime-environment.md).</span></span>

### <a name="initialize-with-officeonready"></a><span data-ttu-id="9e7f9-131">Initialiser avec Office.onReady()</span><span class="sxs-lookup"><span data-stu-id="9e7f9-131">Initialize with Office.onReady()</span></span>

<span data-ttu-id="9e7f9-132">`Office.onReady()` est une méthode asynchrone qui renvoie un objet Promise tandis qu’il vérifie si la bibliothèque Office.js est chargée.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-132">`Office.onReady()` is an asynchronous method that returns a Promise object while it checks to see if the Office.js library is loaded.</span></span> <span data-ttu-id="9e7f9-133">Uniquement lorsque la bibliothèque est chargée, cela résout la promesse sous forme d’objet qui spécifie l’application Office hôte avec une`Office.HostType` valeur enum (`Excel`, `Word`, etc.) et la plateforme avec une`Office.PlatformType` valeur enum (`PC`, `Mac`, `OfficeOnline`, etc..).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-133">When the library is loaded, it resolves the Promise as an object that specifies the Office host application with an `Office.HostType` enum value (`Excel`, `Word`, etc.) and the platform with an `Office.PlatformType` enum value (`PC`, `Mac`, `OfficeOnline`, etc.).</span></span> <span data-ttu-id="9e7f9-134">L’objet Promise se résout immédiatement si la bibliothèque est déjà chargée quand `Office.onReady()` est appelée.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-134">The Promise resolves immediately if the library is already loaded when `Office.onReady()` is called.</span></span>

<span data-ttu-id="9e7f9-135">Une méthode pour appeler `Office.onReady()` consiste à transmettre une méthode de rappel.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-135">One way to call `Office.onReady()` is to pass it a callback method.</span></span> <span data-ttu-id="9e7f9-136">Voici un exemple :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-136">Here's an example:</span></span>

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

<span data-ttu-id="9e7f9-137">Par ailleurs, vous pouvez mettre en chaîne une`then()` méthode permettant d’appeler `Office.onReady()`, au lieu de spécifier un rappel.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-137">Alternatively, you can chain a `then()` method to the call of `Office.onReady()`, instead of passing a callback.</span></span> <span data-ttu-id="9e7f9-138">Par exemple, le code suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-138">For example, the following code checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.onReady()
    .then(function() {
        if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
            console.log("Sorry, this add-in only works with newer versions of Excel.");
        }
    });
```

<span data-ttu-id="9e7f9-139">Voici le même exemple utilisant les mots clés `async` et `await` dans TypeScript :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-139">Here is the same example using the `async` and `await` keywords in TypeScript:</span></span>

```typescript
(async () => {
    await Office.onReady();
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
})();
```

<span data-ttu-id="9e7f9-140">Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être*habituellement* placées dans la réponse à`Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-140">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the response to `Office.onReady()`.</span></span> <span data-ttu-id="9e7f9-141">Par exemple, la fonction `$(document).ready()` de [JQuery](https://jquery.com) sera référencée comme suit :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-141">For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.onReady(function() {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
});
```

<span data-ttu-id="9e7f9-142">Toutefois, il existe des exceptions à cette pratique.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-142">However, there are exceptions to this practice.</span></span> <span data-ttu-id="9e7f9-143">Par exemple, supposons que vous voulez ouvrir votre complément dans un navigateur (au lieu de le charger dans un hôte Office) afin de déboguer votre interface utilisateur avec les outils de navigateur.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-143">For example, suppose you want to open your add-in in a browser (instead of sideload it in an Office host) in order to debug your UI with browser tools.</span></span> <span data-ttu-id="9e7f9-144">Étant donné que Office.js ne sera pas chargé dans le navigateur, `onReady` ne s’exécutera pas et le `$(document).ready` ne s’exécutera pas si cette opération est appelée à l’intérieur d’Office `onReady`.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-144">Since Office.js won't load in the browser, `onReady` won't run and the `$(document).ready` won't run if it's called inside the Office `onReady`.</span></span> <span data-ttu-id="9e7f9-145">Une autre exception : vous souhaitez qu’un indicateur de progression s’affiche dans le volet Office tandis que le complément se charge.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-145">Another exception: you want a progress indicator to appear in the task pane while the add-in is loading.</span></span> <span data-ttu-id="9e7f9-146">Dans ce scénario, votre code doit appeler la jQuery `ready` et utiliser le rappel pour afficher l’indicateur de progression.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-146">In this scenario, your code should call the jQuery `ready` and use it's callback to render the progress indicator.</span></span> <span data-ttu-id="9e7f9-147">Puis le rappel `onReady` Office peut remplacer l’indicateur de progression par l’interface utilisateur final.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-147">Then the Office `onReady`'s callback can replace the progress indicator with the final UI.</span></span> 

### <a name="initialize-with-officeinitialize"></a><span data-ttu-id="9e7f9-148">Initialiser avec Office.initialize</span><span class="sxs-lookup"><span data-stu-id="9e7f9-148">Initialize with Office.initialize</span></span>

<span data-ttu-id="9e7f9-149">Un événement initialisé se déclenche lorsque la bibliothèque Office.js est chargée et prête pour une interaction avec l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-149">An initialize event fires when the Office.js library is loaded and ready for user interaction.</span></span> <span data-ttu-id="9e7f9-150">Vous pouvez attribuer un gestionnaire à `Office.initialize` qui implémente votre logique d’initialisation.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-150">You can assign a handler to `Office.initialize` that implements your initialization logic.</span></span> <span data-ttu-id="9e7f9-151">L’exemple suivant vérifie que la version de l’utilisateur d’Excel prend en charge tous les API que le complément peut appeler.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-151">The following is an example that checks to see that the user's version of Excel supports all the APIs that the add-in might call.</span></span>

```js
Office.initialize = function () {
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.7')) {
        console.log("Sorry, this add-in only works with newer versions of Excel.");
    }
};
```

<span data-ttu-id="9e7f9-152">Si vous utilisez des infrastructures JavaScript supplémentaires qui incluent leur propre gestionnaire d’initialisation ou tests, elles doivent être*habituellement* placées dans l’événement`Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-152">If you are using additional JavaScript frameworks that include their own initialization handler or tests, these should *usually* be placed within the `Office.initialize` event.</span></span> <span data-ttu-id="9e7f9-153">(Mais les exceptions décrites dans la section **initialiser avec Office.onReady()** précédente s’appliquent dans ce cas également.) Par exemple, la fonction[ JQuery](https://jquery.com) `$(document).ready()` serait référencée comme suit :</span><span class="sxs-lookup"><span data-stu-id="9e7f9-153">(But the exceptions described in the **Initialize with Office.onReady()** section earlier apply in this case also.) For example, [JQuery's](https://jquery.com) `$(document).ready()` function would be referenced as follows:</span></span>

```js
Office.initialize = function () {
    // Office is ready
    $(document).ready(function () {
        // The document is ready
    });
  };
```

<span data-ttu-id="9e7f9-154">Pour les compléments de tâches et de contenu, `Office.initialize` fournit un paramètre_raison_ supplémentaire.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-154">For task pane and content add-ins, `Office.initialize` provides an additional _reason_ parameter.</span></span> <span data-ttu-id="9e7f9-155">Ce paramètre peut être utilisé pour savoir comment un complément a été ajouté au document actif.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-155">This parameter specifies how an add-in was added to the current document.</span></span> <span data-ttu-id="9e7f9-156">Vous pouvez l’utiliser pour fournir une logique différente quand un complément est inséré pour la première fois par opposition au moment où il fait déjà partie du document.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-156">You can use this to provide different logic for when an add-in is first inserted versus when it already existed within the document.</span></span>

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

<span data-ttu-id="9e7f9-157">Pour plus d’informations, consultez les pages relatives à l’[événement Office.initialize](/javascript/api/office) et à l’[énumération InitializationReason](/javascript/api/office/office.initializationreason).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-157">For more information, see [Office.initialize Event](/javascript/api/office) and [InitializationReason Enumeration](/javascript/api/office/office.initializationreason).</span></span>

### <a name="major-differences-between-officeinitialize-and-officeonready"></a><span data-ttu-id="9e7f9-158">Principales différences entre Office.initialize et Office.onReady</span><span class="sxs-lookup"><span data-stu-id="9e7f9-158">Major differences between Office.initialize and Office.onReady</span></span>

- <span data-ttu-id="9e7f9-159">Vous ne pouvez assigner qu’un seul gestionnaire à `Office.initialize` et il n’est appelé qu’une seule fois par l’infrastructure d’Office, mais vous pouvez appeler `Office.onReady()`à plusieurs endroits dans votre code et utiliser des rappels différents.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-159">You can assign only one handler to `Office.initialize` and it's called only once by the Office infrastructure; but you can call `Office.onReady()` in different places in your code and use different callbacks.</span></span> <span data-ttu-id="9e7f9-160">Par exemple, votre code pourrait appeler `Office.onReady()` dès que votre script personnalisé charge avec un rappel qui exécute la logique d’initialisation ; et votre code peut également comporter un bouton dans le volet Office dont le script appelle `Office.onReady()` avec un rappel différent.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-160">For example, your code could call `Office.onReady()` as soon as your custom script loads with a callback that runs initialization logic; and your code could also have a button in the task pane, whose script calls `Office.onReady()` with a different callback.</span></span> <span data-ttu-id="9e7f9-161">Si c’est le cas, le deuxième rappel s’exécute quand l’utilisateur clique sur le bouton.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-161">If so, the second callback runs when the button is clicked.</span></span>

- <span data-ttu-id="9e7f9-162">L’événement`Office.initialize` se déclenche à la fin du processus interne dans lequel Office.js s’initialise lui-même.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-162">The `Office.initialize` event fires at the end of the internal process in which Office.js initializes itself.</span></span> <span data-ttu-id="9e7f9-163">Et il se déclenche *immédiatement* après la fin du processus interne.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-163">And it fires *immediately* after the internal process ends.</span></span> <span data-ttu-id="9e7f9-164">Si le code dans lequel vous attribuez un gestionnaire à l’événement s’exécute trop longtemps après le déclenchement de l’événement, votre gestionnaire ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-164">If the code in which you assign a handler to the event executes too long after the event fires, then your handler doesn't run.</span></span> <span data-ttu-id="9e7f9-165">Par exemple, si vous utilisez le Gestionnaire des tâches WebPack, il peut configurer page d’accueil du complément pour charger les fichiers polyfill une fois que le serveur charge Office.js mais avant que le serveur ne charge votre code JavaScript personnalisé.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-165">For example, if you are using the WebPack task manager, it might configure the add-in's home page to load polyfill files after it loads Office.js but before it loads your custom JavaScript.</span></span> <span data-ttu-id="9e7f9-166">Le temps que votre script se charge et affecte le Gestionnaire, l’événement initialisé s’est déjà produit.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-166">By the time your script loads and assigns the handler, the initialize event has already happened.</span></span> <span data-ttu-id="9e7f9-167">Mais il n’est jamais « trop tard » pour appeler `Office.onReady()`.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-167">But it is never "too late" to call `Office.onReady()`.</span></span> <span data-ttu-id="9e7f9-168">Si l’événement initialisé s’est déjà produit, le rappel s’exécute immédiatement.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-168">If the initialize event has already happened, the callback runs immediately.</span></span>

> [!NOTE]
> <span data-ttu-id="9e7f9-169">Même si vous n’avez aucune logique de démarrage, appelez `Office.onReady()` ou attribuez une fonction vide à `Office.initialize` lorsque votre complément JavaScript se charge.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-169">Even if you have no start-up logic, you should either call `Office.onReady()` or assign an empty function to `Office.initialize` when your add-in JavaScript loads.</span></span> <span data-ttu-id="9e7f9-170">Certaines combinaisons de plateforme et d’hôte Office ne chargeront pas le volet Office tant que l’une de ces situations se produisent.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-170">Some Office host and platform combinations won't load the task pane until one of these happens.</span></span> <span data-ttu-id="9e7f9-171">Les exemples suivants présentent ces deux approches.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-171">The following examples show these two approaches.</span></span>
>
>```js  
>Office.onReady();
>```
>
>
>```js
>Office.initialize = function () {};
>```

## <a name="office-javascript-api-object-model"></a><span data-ttu-id="9e7f9-172">Modèle d’objet API JavaScript Office</span><span class="sxs-lookup"><span data-stu-id="9e7f9-172">Office JavaScript API object model</span></span>

<span data-ttu-id="9e7f9-173">Une fois initialisé, le complément peut interagir avec l’hôte (par exemple, Excel, Outlook).</span><span class="sxs-lookup"><span data-stu-id="9e7f9-173">Once initialized, the add-in can interact with the host (e.g. Excel, Outlook).</span></span> <span data-ttu-id="9e7f9-174">La page [Modèle objet API JavaScript Office](office-javascript-api-object-model.md) comporte plus d’informations sur les modèles d’utilisation spécifiques.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-174">The [Office JavaScript API object model](office-javascript-api-object-model.md) page has more details on specific usage patterns.</span></span> <span data-ttu-id="9e7f9-175">Il existe également une documentation de référence détaillée pour les deux[ APIs Communes](/office/dev/add-ins/reference/javascript-api-for-office) et spécifiques.</span><span class="sxs-lookup"><span data-stu-id="9e7f9-175">There is also detailed reference documentation for both [Common APIs](/office/dev/add-ins/reference/javascript-api-for-office) and host-specific APIs.</span></span>
