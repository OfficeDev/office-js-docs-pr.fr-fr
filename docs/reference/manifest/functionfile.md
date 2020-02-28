---
title: Élément FunctionFile dans le fichier manifest
description: ''
ms.date: 10/09/2018
localization_priority: Normal
ms.openlocfilehash: eec1dc8eb2e099670469af6ef300592fc4a31e64
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324868"
---
# <a name="functionfile-element"></a><span data-ttu-id="a230c-102">Élément FunctionFile</span><span class="sxs-lookup"><span data-stu-id="a230c-102">FunctionFile element</span></span>

<span data-ttu-id="a230c-103">Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a230c-103">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI.</span></span> <span data-ttu-id="a230c-104">L' `FunctionFile` élément est un élément enfant de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="a230c-104">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="a230c-105">L' `resid` attribut de l' `FunctionFile` élément est défini sur la valeur de l' `id` attribut d’un `Url` élément dans l' `Resources` élément qui contient l’URL d’un fichier HTML qui contient ou charge toutes les fonctions JavaScript utilisées par les boutons de commande de complément sans interface utilisateur, comme défini par l' [élément Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="a230c-105">The `resid` attribute of the `FunctionFile` element is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="a230c-106">Voici un exemple de l' `FunctionFile` élément.</span><span class="sxs-lookup"><span data-stu-id="a230c-106">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="a230c-107">Le code JavaScript dans le fichier HTML indiqué par `FunctionFile` l’élément doit `Office.initialize` appeler et définir des fonctions nommées qui prennent un `event`seul paramètre :.</span><span class="sxs-lookup"><span data-stu-id="a230c-107">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="a230c-108">Les fonctions doivent utiliser l’`item.notificationMessages` API pour indiquer l’avancement, réussite ou Échec de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a230c-108">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="a230c-109">Elle doit également appeler `event.completed` lorsqu’il a fini d’exécution.</span><span class="sxs-lookup"><span data-stu-id="a230c-109">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="a230c-110">Le nom des fonctions est utilisé dans l' `FunctionName` élément pour les boutons sans interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="a230c-110">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="a230c-111">Voici un exemple de fichier HTML définissant une `trackMessage` fonction.</span><span class="sxs-lookup"><span data-stu-id="a230c-111">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

```js
Office.initialize = function () {
    doAuth();
}

function trackMessage (event) {
    var buttonId = event.source.id;    
    var itemId = Office.context.mailbox.item.id;
    // save this message
    event.completed();
}
```

<span data-ttu-id="a230c-112">Le code suivant montre comment implémenter la fonction utilisée par `FunctionName`.</span><span class="sxs-lookup"><span data-stu-id="a230c-112">The following code shows how to implement the function used by `FunctionName`.</span></span>

```js
// The initialize function must be run each time a new page is loaded.
(function () {
    Office.initialize = function (reason) {
        // If you need to initialize something you can do so here.
    };
})();

// Your function must be in the global namespace.
function writeText(event) {

    // Implement your custom code here. The following code is a simple example.

    Office.context.document.setSelectedDataAsync("ExecuteFunction works. Button ID=" + event.source.id,
        function (asyncResult) {
            var error = asyncResult.error;
            if (asyncResult.status === "failed") {
                // Show error message.
            }
            else {
                // Show success message.
            }
        });
    // Calling event.completed is required. event.completed lets the platform know that processing has completed.
    event.completed();
}
```

> [!IMPORTANT]
> <span data-ttu-id="a230c-113">L’appel de `event.completed` signale que vous avez réussi à gérer l’événement.</span><span class="sxs-lookup"><span data-stu-id="a230c-113">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="a230c-114">Lorsqu’une fonction est appelée à plusieurs reprises, par exemple lorsque l’utilisateur clique plusieurs fois sur une même commande de complément, tous les événements sont automatiquement mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="a230c-114">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="a230c-115">Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente.</span><span class="sxs-lookup"><span data-stu-id="a230c-115">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="a230c-116">Lorsque votre fonction appelle `event.completed`, l’appel suivant de cette fonction s’exécute.</span><span class="sxs-lookup"><span data-stu-id="a230c-116">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="a230c-117">Vous devez appeler `event.completed`; Sinon, votre fonction ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="a230c-117">You must call `event.completed`; otherwise your function will not run.</span></span>
