---
title: Élément FunctionFile dans le fichier manifest
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: 634d383498698b55990dc73e66ec11616396f968
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: fr-FR
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432696"
---
# <a name="functionfile-element"></a><span data-ttu-id="7c70c-102">Élément FunctionFile</span><span class="sxs-lookup"><span data-stu-id="7c70c-102">FunctionFile element</span></span>

<span data-ttu-id="7c70c-p101">Spécifie le fichier de code source pour les opérations qu’un complément expose via les commandes de complément qui exécutent une fonction JavaScript au lieu d’afficher l’interface utilisateur. L’élément **FunctionFile** est un élément enfant de [DesktopFormFactor](desktopformfactor.md) ou de [MobileFormFactor](mobileformfactor.md). L’attribut **resid** de l’élément **FunctionFile** est défini sur la valeur de l’attribut **id** d’un élément **Url** dans l’élément **Resources** contenant l’URL d’un fichier HTML qui contient ou charge toutes les fonctions JavaScript utilisées par les boutons de commande de complément sans interface utilisateur, telles que définies par l’élément [Control](control.md).</span><span class="sxs-lookup"><span data-stu-id="7c70c-p101">Specifies the source code file for operations that an add-in exposes through add-in commands that execute a JavaScript function instead of displaying UI. The  **FunctionFile** element is a child element of [DesktopFormFactor](desktopformfactor.md) or [MobileFormFactor](mobileformfactor.md). The **resid** attribute of the **FunctionFile** element is set to the value of the **id** attribute of a **Url** element in the **Resources** element that contains the URL to an HTML file that contains or loads all  the JavaScript functions used by UI-less add-in command buttons, as defined by the [Control element](control.md).</span></span>

<span data-ttu-id="7c70c-106">Vous trouverez ci-dessous un exemple de l’élément **FunctionFile**.</span><span class="sxs-lookup"><span data-stu-id="7c70c-106">The following is an example of the **FunctionFile** element.</span></span>

```XML
<DesktopFormFactor>
  <FunctionFile resid="residDesktopFuncUrl" />
  <ExtensionPoint xsi:type="PrimaryCommandSurface">
    <!-- information about this extension point -->
  </ExtensionPoint>

  <!-- You can define more than one ExtensionPoint element as needed -->

</DesktopFormFactor>
```

<span data-ttu-id="7c70c-107">JavaScript dans le fichier HTML indiqué par l’élément**FunctionFile**doit appeler`Office.initialize`et définir nommées fonctions qui acceptent un paramètre unique: `event`.</span><span class="sxs-lookup"><span data-stu-id="7c70c-107">The JavaScript in the HTML file indicated by the  **FunctionFile** element must call `Office.initialize` and define named functions that take a single parameter: `event`.</span></span> <span data-ttu-id="7c70c-108">Les fonctions doivent utiliser l’`item.notificationMessages` API pour indiquer l’avancement, réussite ou Échec de l’utilisateur.</span><span class="sxs-lookup"><span data-stu-id="7c70c-108">The functions should use the `item.notificationMessages` API to indicate progress, success, or failure to the user.</span></span> <span data-ttu-id="7c70c-109">Elle doit également appeler `event.completed` lorsqu’il a fini d’exécution.</span><span class="sxs-lookup"><span data-stu-id="7c70c-109">It should also call `event.completed` when it has finished execution.</span></span> <span data-ttu-id="7c70c-110">Le nom des fonctions sont utilisés dans le **FunctionName** l’élément de l’interface utilisateur moins boutons.</span><span class="sxs-lookup"><span data-stu-id="7c70c-110">The name of the functions are used in the **FunctionName** element for UI-less buttons.</span></span>

<span data-ttu-id="7c70c-111">Vous trouverez ci-dessous un exemple d’un fichier HTML définissant une fonction **trackMessage**.</span><span class="sxs-lookup"><span data-stu-id="7c70c-111">The following is an example of an HTML file defining a **trackMessage** function.</span></span>

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

<span data-ttu-id="7c70c-112">Le code suivant montre comment implémenter la fonction utilisée par**FunctionName**.</span><span class="sxs-lookup"><span data-stu-id="7c70c-112">The following code shows how to implement the function used by **FunctionName**.</span></span>

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
> <span data-ttu-id="7c70c-113">L’appel de l’élément**event.completed**indique que vous avez correctement géré l’événement.</span><span class="sxs-lookup"><span data-stu-id="7c70c-113">IMPORTANT  The call to **event.completed** signals that you have successfully handled the event.</span></span> <span data-ttu-id="7c70c-114">Lorsqu’une fonction est appelée à plusieurs reprises, par exemple, lorsque l’utilisateur clique plusieurs fois sur une même commande de complément, tous les événements sont automatiquement mis en file d’attente.</span><span class="sxs-lookup"><span data-stu-id="7c70c-114">When a function is called multiple times, such as multiple clicks on the same add-in command, all events are automatically queued.</span></span> <span data-ttu-id="7c70c-115">Le premier événement s’exécute automatiquement, tandis que les autres événements restent dans la file d’attente.</span><span class="sxs-lookup"><span data-stu-id="7c70c-115">The first event runs automatically, while the other events remain on the queue.</span></span> <span data-ttu-id="7c70c-116">Lorsque votre fonction appelle**event.completed**, l’appel suivant de cette fonction s’exécute.</span><span class="sxs-lookup"><span data-stu-id="7c70c-116">When your function calls **event.completed**, the next queued call to that function runs.</span></span> <span data-ttu-id="7c70c-117">Vous devez appeler**event.completed** pour que votre fonction s’exécute correctement.</span><span class="sxs-lookup"><span data-stu-id="7c70c-117">You must implement **event.completed**, otherwise your function will not run.</span></span>