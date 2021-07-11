---
title: Élément FunctionFile dans le fichier manifest
description: Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur.
ms.date: 11/06/2020
localization_priority: Normal
ms.openlocfilehash: f31a1bc7a561305a89f5388102a4985aaa31fe37
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: fr-FR
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348299"
---
# <a name="functionfile-element"></a><span data-ttu-id="d0b80-103">Élément FunctionFile</span><span class="sxs-lookup"><span data-stu-id="d0b80-103">FunctionFile element</span></span>

<span data-ttu-id="d0b80-104">Spécifie le fichier de code source pour les opérations qu’un add-in expose de l’une des manières suivantes.</span><span class="sxs-lookup"><span data-stu-id="d0b80-104">Specifies the source code file for operations that an add-in exposes in one of the following ways.</span></span>

* <span data-ttu-id="d0b80-105">Commandes de add-in qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d0b80-105">Add-in commands that execute a JavaScript function instead of displaying UI.</span></span>
* <span data-ttu-id="d0b80-106">Raccourcis clavier qui exécutent une fonction JavaScript.</span><span class="sxs-lookup"><span data-stu-id="d0b80-106">Keyboard shortcuts that execute a JavaScript function.</span></span>

<span data-ttu-id="d0b80-107">`FunctionFile`L’élément est un élément enfant de [DesktopFormFactor](desktopformfactor.md) ou [MobileFormFactor](mobileformfactor.md).</span><span class="sxs-lookup"><span data-stu-id="d0b80-107">The `FunctionFile` element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md).</span></span> <span data-ttu-id="d0b80-108">L’attribut de l’élément ne peut pas être plus de 32 caractères et est défini sur la valeur de l’attribut d’un élément dans l’élément qui contient l’URL d’un fichier HTML qui contient ou charge toutes les fonctions JavaScript utilisées par les boutons de commande de l’interface utilisateur sans interface `resid` `FunctionFile` `id` `Url` `Resources` utilisateur, [](control.md)comme défini par l’élément Control .</span><span class="sxs-lookup"><span data-stu-id="d0b80-108">The `resid` attribute of the `FunctionFile` element can be no more than 32 characters and is set to the value of the `id` attribute of a `Url` element in the `Resources` element that contains the URL to an HTML file that contains or loads all the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="d0b80-109">Voici un exemple de `FunctionFile` l’élément.</span><span class="sxs-lookup"><span data-stu-id="d0b80-109">The following is an example of the `FunctionFile` element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="d0b80-110">Le code JavaScript dans le fichier HTML indiqué par l’élément doit appeler et définir des fonctions nommées qui `FunctionFile` `Office.initialize` prennent un seul paramètre : `event` .</span><span class="sxs-lookup"><span data-stu-id="d0b80-110">The JavaScript in the HTML file indicated by the `FunctionFile` element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="d0b80-111">Les fonctions doivent utiliser l’`item.notificationMessages` API pour indiquer l’avancement, réussite ou Échec de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d0b80-111">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="d0b80-112">Elle doit également appeler `event.completed` lorsqu’il a fini d’exécution.</span><span class="sxs-lookup"><span data-stu-id="d0b80-112">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="d0b80-113">Le nom des fonctions est utilisé dans l’élément pour les boutons sans interface `FunctionName` utilisateur.</span><span class="sxs-lookup"><span data-stu-id="d0b80-113">The name of the functions are used in the `FunctionName` element for UI-less buttons.</span></span>

<span data-ttu-id="d0b80-114">Voici un exemple de fichier HTML définissant une `trackMessage` fonction.</span><span class="sxs-lookup"><span data-stu-id="d0b80-114">The following is an example of an HTML file defining a `trackMessage` function.</span></span>

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

<span data-ttu-id="d0b80-115">Le code suivant montre comment implémenter la fonction utilisée par `FunctionName` .</span><span class="sxs-lookup"><span data-stu-id="d0b80-115">The following code shows how to implement the function used by `FunctionName`.</span></span>

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
> <span data-ttu-id="d0b80-116">L’appel `event.completed` aux signaux que vous avez géré l’événement avec succès.</span><span class="sxs-lookup"><span data-stu-id="d0b80-116">The call to `event.completed` signals that you have successfully handled the event.</span></span> <span data-ttu-id="d0b80-117">Lorsqu’une fonction est appelée à plusieurs reprises, par exemple lorsque l’utilisateur clique plusieurs fois sur une même commande de complément, tous les événements sont automatiquement mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="d0b80-117">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="d0b80-118">Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente.</span><span class="sxs-lookup"><span data-stu-id="d0b80-118">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="d0b80-119">Lorsque votre fonction appelle, l’appel mis en file `event.completed` d’attente suivant à cette fonction s’exécute.</span><span class="sxs-lookup"><span data-stu-id="d0b80-119">When your function calls `event.completed`, the next queued call to that function runs.</span></span> <span data-ttu-id="d0b80-120">Vous devez appeler `event.completed` ; sinon, votre fonction ne s’exécutera pas.</span><span class="sxs-lookup"><span data-stu-id="d0b80-120">You must call `event.completed`; otherwise your function will not run.</span></span>
